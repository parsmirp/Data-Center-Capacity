# ============================================================
# Weekly DC Capacity Report (email to the sales team)
#
# Put this file in the same folder as app.R (~/Desktop/shiny).
# Run it in RStudio with "Source", or from Terminal:
#   Rscript ~/Desktop/shiny/weekly_report.R
#
# What it does:
#   1. Reuses the parsing functions from app.R (same numbers as the dashboard)
#   2. Reads DC.xlsx and compares it with last week's data
#   3. Builds an HTML email and saves a preview: weekly_report_preview.html
#   4. Sends it through your Outlook / Microsoft 365 account
#
# TEST_MODE = TRUE  -> sends only to you and does NOT move the
#                      "last week" baseline, so you can re-run it freely.
# TEST_MODE = FALSE -> sends to the sales team and saves this week's
#                      data as the baseline for next week.
#
# One-time setup:
#   install.packages("Microsoft365R")
#   The first send opens a browser window to sign in to Microsoft 365.
# ============================================================

# ---------- Settings ----------

APP_DIR      <- "~/Desktop/shiny"
TEST_MODE    <- TRUE
SEND_EMAIL   <- FALSE          # FALSE = only build the preview file, don't send

MY_EMAIL     <- "parsa.mirpour@bytebt.com"
SALES_TEAM   <- c("parsa.mirpour@bytebt.com",
                  "glen.inocencio@bytebt.com",
                  "ray@bytebt.com", 
                  "karen.lau@bytebt.com",
                  "jacob.gomez-decena@bytebt.com",
                  "kage.russell@bytebt.com")

DASHBOARD_URL         <- "https://bytebt.shinyapps.io/shiny"   # link to the hosted dashboard (optional)
REPORT_SHAREPOINT_URL <- ""   # link to the master Excel (optional; falls back to app.R's)

LOGO_FILE   <- "www/logo.png"  # same logo the dashboard uses; set to "" to leave it out
LOGO_HEIGHT <- 44              # px in the email header

COMING_SOON_QUARTERS <- 1     # this quarter + the next # quarters
STALE_DAYS           <- 60    # "not updated in X days"
MIN_MW_CHANGE        <- 0.4   # ignore capacity changes smaller than this (MW)
MAX_ROWS             <- 15    # rows shown per table before "+N more"

STATE_DB <- "weekly_report_state.sqlite"   # the report's own memory (baseline + change tracking)

# ---------- Load the helper functions from app.R ----------
# Everything above the "STARTUP IMPORT" section of app.R: packages, quarter
# helpers, capacity parsing, clean_master_df(), database helpers.
# Nothing is imported or geocoded, and the Shiny app is not started.

setwd(path.expand(APP_DIR))

app_lines <- readLines("app.R", warn = FALSE, encoding = "UTF-8")
cut_at <- grep("^# STARTUP IMPORT", app_lines)[1]
if (is.na(cut_at)) stop("Could not find the 'STARTUP IMPORT' section in app.R.")
eval(parse(text = app_lines[seq_len(cut_at - 2)], encoding = "UTF-8"), envir = globalenv())

# ifelse() returns TRUE/FALSE (not text) when it gets zero rows, which breaks
# bind_rows() when one of the tables is empty. Wrap the text helpers so they
# always return text, whatever app.R's version looks like.
txt_ifelse <- function(...) as.character(ifelse(...))
local({
  wrap <- function(f) { force(f); function(...) as.character(f(...)) }
  for (fn in c("fmt_mw", "mw_span_txt", "mw_range_txt", "quarter_label")) {
    assign(fn, wrap(get(fn, envir = globalenv())), envir = globalenv())
  }
})

if (!nzchar(REPORT_SHAREPOINT_URL) && exists("SHAREPOINT_URL")) REPORT_SHAREPOINT_URL <- SHAREPOINT_URL

TODAY <- Sys.Date()

# ============================================================
# SITE SUMMARIES
# A "site" = Operator + Market + State + Country. Rows that share all four
# are added together, so re-ordering or splitting rows in the sheet doesn't
# show up as fake changes.
# ============================================================

norm_txt <- function(x) str_to_lower(str_squish(coalesce(as.character(x), "")))

site_key <- function(op, mkt, st, ctry) {
  paste(norm_txt(op), norm_txt(mkt), norm_txt(st), norm_txt(ctry), sep = "|")
}

summarise_sites <- function(current, pipeline) {
  if (!"City_clean" %in% names(current)) {
    current$City_clean <- coalesce(current$Market %||% NA_character_, current$City)
  }

  sites <- current %>%
    filter(!is.na(Operator)) %>%
    mutate(key = site_key(Operator, City_clean, State, Country),
           Capacity_MW_est = as.numeric(Capacity_MW_est)) %>%
    group_by(key) %>%
    summarise(
      Operator = first(Operator),
      Market   = first(City_clean),
      State    = first(State),
      Country  = first(Country),
      Total_MW = if (all(is.na(Capacity_MW_est))) NA_real_ else sum(Capacity_MW_est, na.rm = TRUE),
      Rows     = n(),        # each sheet row is its own listing (suite, floor, hall...)
      .groups = "drop"
    )

  pipe <- if (nrow(pipeline) == 0) {
    tibble(key = character(), Quarter = character(), Quarter_Date = as.Date(character()), MW = double())
  } else {
    pipeline %>%
      filter(!is.na(Operator)) %>%
      mutate(
        key = site_key(Operator, City_clean, State, Country),
        Quarter = str_to_upper(Quarter),
        Quarter_Date = quarter_to_date(Quarter),
        MW_available = as.numeric(MW_available)
      ) %>%
      filter(!is.na(Quarter_Date), !is.na(MW_available), MW_available > 0) %>%
      group_by(key, Quarter, Quarter_Date) %>%
      summarise(MW = sum(MW_available), .groups = "drop")
  }

  firsts <- pipe %>%
    group_by(key) %>%
    summarise(First_Q_Date = min(Quarter_Date), Peak_MW = max(MW), .groups = "drop")

  sites <- sites %>% left_join(firsts, by = "key")

  list(sites = sites, pipe = pipe)
}

# ============================================================
# THIS WEEK'S DATA (straight from DC.xlsx)
# ============================================================

if (!file.exists(MASTER_FILE)) stop("Could not find ", MASTER_FILE, " in ", getwd())

raw <- readxl::read_excel(MASTER_FILE, col_types = "text")
names(raw) <- str_squish(names(raw))
parsed <- clean_master_df(raw)
now <- summarise_sites(parsed$current, parsed$pipeline)

# "Capacity Upload date" per site (used for the stale check when filled in)
parse_sheet_date <- function(x) {
  x <- str_trim(as.character(x))
  out <- rep(as.Date(NA), length(x))
  num <- suppressWarnings(as.numeric(x))
  is_num <- !is.na(num) & num > 20000 & num < 80000          # Excel date serials
  out[is_num] <- as.Date(num[is_num], origin = "1899-12-30")
  for (f in c("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%b %d, %Y", "%B %d, %Y")) {
    todo <- is.na(out) & !is.na(x)
    if (!any(todo)) break
    out[todo] <- suppressWarnings(as.Date(x[todo], format = f))
  }
  out
}

upload_col <- if ("Capacity Upload date" %in% names(raw)) raw$`Capacity Upload date` else rep(NA_character_, nrow(raw))

upload_dates <- parsed$current %>%
  mutate(Upload_Date = parse_sheet_date(upload_col[row_id])) %>%
  filter(!is.na(Operator), !is.na(Upload_Date)) %>%
  mutate(key = site_key(Operator, City_clean, State, Country)) %>%
  select(key, Upload_Date)

# Only take the max when there are dates (the column is blank in the sheet
# today, and max() of nothing warns and returns -Inf)
if (nrow(upload_dates) > 0) {
  upload_dates <- upload_dates %>%
    group_by(key) %>%
    summarise(Upload_Date = max(Upload_Date), .groups = "drop")
}

# ============================================================
# LAST WEEK'S DATA (the baseline)
# Every real send saves a dated snapshot (the last 12 are kept).
# The baseline is the newest snapshot that is at least 6 days old, so
# running the report twice in one week never compares today with today.
#   1st choice: this script's own snapshot, 6+ days old
#   2nd choice: the app's Version History, 6+ days old
#   Otherwise:  the oldest data available
# ============================================================

MIN_BASELINE_AGE <- 6     # days
KEEP_SNAPSHOTS   <- 12

state_con <- function() dbConnect(RSQLite::SQLite(), STATE_DB)

# Older versions of this script kept a single undated snapshot (base_* tables).
# Move it into the dated snapshot tables once.
migrate_old_baseline <- function(con) {
  tbls <- dbListTables(con)
  if (!all(c("base_sites", "base_pipe", "base_info") %in% tbls)) return(invisible())
  as_of <- dbReadTable(con, "base_info")$as_of[1]
  dbWriteTable(con, "snap_sites", dbReadTable(con, "base_sites") %>% mutate(as_of = as_of),
               append = "snap_sites" %in% tbls)
  dbWriteTable(con, "snap_pipe", dbReadTable(con, "base_pipe") %>% mutate(as_of = as_of),
               append = "snap_pipe" %in% tbls)
  for (t in c("base_sites", "base_pipe", "base_info")) dbRemoveTable(con, t)
}

snapshot_dates <- function() {
  con <- state_con()
  on.exit(dbDisconnect(con), add = TRUE)
  migrate_old_baseline(con)
  if (!"snap_sites" %in% dbListTables(con)) return(as.Date(character()))
  sort(unique(as.Date(dbGetQuery(con, "SELECT DISTINCT as_of FROM snap_sites")$as_of)))
}

load_own_baseline <- function(as_of) {
  con <- state_con()
  on.exit(dbDisconnect(con), add = TRUE)
  d <- as.character(as_of)

  sites <- dbGetQuery(con, "SELECT * FROM snap_sites WHERE as_of = ?", params = list(d)) %>%
    as_tibble() %>%
    mutate(First_Q_Date = quarter_to_date(First_Q)) %>%
    select(-First_Q, -as_of)
  # Snapshots saved before rows were counted can't give accurate counts
  if (!"Rows" %in% names(sites) || all(is.na(sites$Rows))) return(NULL)
  pipe <- dbGetQuery(con, "SELECT * FROM snap_pipe WHERE as_of = ?", params = list(d)) %>%
    as_tibble() %>%
    mutate(Quarter_Date = quarter_to_date(Quarter)) %>%
    select(-as_of)

  list(data = list(sites = sites, pipe = pipe),
       as_of = as_of,
       source = "previous report")
}

save_own_baseline <- function(summary) {
  con <- state_con()
  on.exit(dbDisconnect(con), add = TRUE)
  migrate_old_baseline(con)
  d <- as.character(TODAY)
  # Older snapshot tables have no row counts: start them fresh
  if ("snap_sites" %in% dbListTables(con) && !"Rows" %in% dbListFields(con, "snap_sites")) {
    for (t in intersect(c("snap_sites", "snap_pipe"), dbListTables(con))) dbRemoveTable(con, t)
  }
  tbls <- dbListTables(con)

  # Re-sending on the same day replaces that day's snapshot
  for (t in intersect(c("snap_sites", "snap_pipe"), tbls)) {
    dbExecute(con, sprintf("DELETE FROM %s WHERE as_of = ?", t), params = list(d))
  }
  dbWriteTable(con, "snap_sites",
               summary$sites %>% mutate(First_Q = quarter_label(First_Q_Date), as_of = d) %>% select(-First_Q_Date),
               append = "snap_sites" %in% tbls)
  dbWriteTable(con, "snap_pipe",
               summary$pipe %>% transmute(key, Quarter, MW, as_of = d),
               append = "snap_pipe" %in% tbls)

  # Keep only the most recent snapshots
  keep <- dbGetQuery(con, sprintf(
    "SELECT DISTINCT as_of FROM snap_sites ORDER BY as_of DESC LIMIT %d", KEEP_SNAPSHOTS))$as_of
  marks <- paste(rep("?", length(keep)), collapse = ",")
  for (t in c("snap_sites", "snap_pipe")) {
    dbExecute(con, sprintf("DELETE FROM %s WHERE as_of NOT IN (%s)", t, marks), params = as.list(keep))
  }
}

load_app_baseline <- function(min_age_days = MIN_BASELINE_AGE, require_old = FALSE) {
  if (!file.exists(DB_PATH)) return(NULL)
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  tbls <- dbListTables(con)
  if (!"dc_meta" %in% tbls) return(NULL)

  meta <- dbReadTable(con, "dc_meta") %>%
    as_tibble() %>%
    mutate(ts = suppressWarnings(as.POSIXct(timestamp))) %>%
    filter(!is.na(ts), paste0("dc_history_v", version) %in% tbls)
  if (nrow(meta) == 0) return(NULL)

  old <- meta %>% filter(ts <= Sys.time() - min_age_days * 86400)
  if (nrow(old) == 0 && require_old) return(NULL)
  pick <- if (nrow(old) > 0) slice_max(old, version, n = 1) else slice_min(meta, version, n = 1)

  cur <- dbReadTable(con, paste0("dc_history_v", pick$version)) %>% as_tibble()
  pipe_tbl <- paste0("dc_pipeline_history_v", pick$version)
  pipe <- if (pipe_tbl %in% tbls) dbReadTable(con, pipe_tbl) %>% as_tibble() else tibble()
  if (nrow(pipe) == 0) {
    pipe <- tibble(Operator = character(), City_clean = character(), State = character(),
                   Country = character(), Quarter = character(), MW_available = double())
  }

  list(data = summarise_sites(cur, pipe),
       as_of = as.Date(pick$ts),
       source = sprintf("dashboard version %d", pick$version))
}

snaps     <- snapshot_dates()
old_snaps <- snaps[snaps <= TODAY - MIN_BASELINE_AGE]

baseline <- NULL
for (d in rev(as.character(old_snaps))) {            # newest usable snapshot first
  baseline <- load_own_baseline(as.Date(d))
  if (!is.null(baseline)) break
}
if (is.null(baseline)) baseline <- load_app_baseline(require_old = TRUE)
if (is.null(baseline)) baseline <- load_app_baseline()
for (d in as.character(snaps)) {                      # last resort: oldest usable snapshot
  if (!is.null(baseline)) break
  baseline <- load_own_baseline(as.Date(d))
}

# ============================================================
# RENAMED OPERATORS
# A site that disappeared and a new site at the same Market/State/Country
# whose operator names match (one contains the other, ignoring words like
# "Data Centers", "Inc", "LLC") are treated as the SAME site, renamed.
#   "Vantage Data Centers" -> "Vantage"   = rename
#   "Equinix" -> "Digital Realty"         = really removed + really new
# ============================================================

core_name <- function(x) {
  x <- norm_txt(x)
  x <- str_replace_all(x, "[[:punct:]]", " ")
  x <- str_remove_all(x, "\\b(data ?cent(er|re)s?|dcs?|inc|llc|ltd|corp|corporation|co|company|group|holdings)\\b")
  str_squish(x)
}

loc_key <- function(df) paste(norm_txt(df$Market), norm_txt(df$State), norm_txt(df$Country), sep = "|")

find_renames <- function(ns, bs) {
  none <- tibble(old_key = character(), new_key = character(),
                 old_op = character(), new_op = character())
  gone <- bs %>% anti_join(ns, by = "key")
  new  <- ns %>% anti_join(bs, by = "key")
  if (nrow(gone) == 0 || nrow(new) == 0) return(none)

  cand <- inner_join(
    gone %>% transmute(old_key = key, loc = loc_key(gone), old_core = core_name(Operator), old_op = Operator),
    new  %>% transmute(new_key = key, loc = loc_key(new),  new_core = core_name(Operator), new_op = Operator),
    by = "loc", relationship = "many-to-many"
  ) %>%
    filter(
      nchar(old_core) >= 3, nchar(new_core) >= 3,
      old_core == new_core |
        str_detect(new_core, fixed(old_core)) |
        str_detect(old_core, fixed(new_core))
    )

  # Only clear one-to-one matches; anything ambiguous stays as removed/new
  cand %>%
    add_count(old_key, name = "n_old") %>%
    add_count(new_key, name = "n_new") %>%
    filter(n_old == 1, n_new == 1) %>%
    select(old_key, new_key, old_op, new_op)
}

renames <- if (is.null(baseline)) find_renames(now$sites, now$sites[0, ]) else
  find_renames(now$sites, baseline$data$sites)

if (nrow(renames) > 0) {
  # Point last week's rows at the new names so every comparison treats them as the same site
  remap <- function(keys) {
    coalesce(renames$new_key[match(keys, renames$old_key)], keys)
  }
  baseline$data$sites$key <- remap(baseline$data$sites$key)
  baseline$data$pipe$key  <- remap(baseline$data$pipe$key)
  cat(sprintf("Detected %d renamed site(s); compared them as the same sites.\n", nrow(renames)))
}

rename_summary <- renames %>%
  count(old_op, new_op, name = "n_sites") %>%
  mutate(txt = sprintf("%s → %s (%d site%s)", old_op, new_op, n_sites, ifelse(n_sites == 1, "", "s")))

# ============================================================
# CHANGE TRACKING (for the stale check)
# Remembers when each site's numbers last changed. The sheet's
# "Capacity Upload date" wins whenever it's newer.
# ============================================================

site_signature <- function(summary) {
  q <- summary$pipe %>%
    arrange(key, Quarter_Date) %>%
    group_by(key) %>%
    summarise(q = paste0(Quarter, "=", round(MW, 2), collapse = ";"), .groups = "drop")
  summary$sites %>%
    left_join(q, by = "key") %>%
    transmute(key, sig = paste(round(coalesce(Total_MW, -1), 2), coalesce(q, ""), sep = "#"))
}

update_site_history <- function(summary, renames) {
  con <- state_con()
  on.exit(dbDisconnect(con), add = TRUE)

  sig_now <- site_signature(summary)

  if (!"site_history" %in% dbListTables(con)) {
    # First run: we don't know when sites last changed, so leave it blank
    hist <- sig_now %>% mutate(first_seen = as.character(TODAY), last_changed = NA_character_)
  } else {
    old <- dbReadTable(con, "site_history") %>% as_tibble()
    # Renamed sites keep their history instead of starting over as "new"
    if (nrow(renames) > 0) {
      i <- match(old$key, renames$old_key)
      was_renamed <- !is.na(i)
      old$key[was_renamed] <- renames$new_key[i[was_renamed]]
      # If the new name was already recorded by an earlier test run, keep the older history
      old <- old[order(!was_renamed), ] %>% distinct(key, .keep_all = TRUE)
    }
    hist <- sig_now %>%
      left_join(old %>% rename(old_sig = sig), by = "key") %>%
      mutate(
        last_changed = case_when(
          is.na(first_seen) ~ as.character(TODAY),          # brand-new site
          old_sig != sig    ~ as.character(TODAY),          # numbers changed
          TRUE              ~ last_changed
        ),
        first_seen = coalesce(first_seen, as.character(TODAY))
      ) %>%
      select(key, sig, first_seen, last_changed)
  }

  dbWriteTable(con, "site_history", hist, overwrite = TRUE)
  hist
}

site_hist <- update_site_history(now, renames)

# ============================================================
# BUILD THE SECTIONS
# ============================================================

# ---- Colours & cell helpers (everything below builds safe HTML cells) ----

C_NAVY <- "#0B2545"; C_BLUE <- "#2563EB"; C_TEXT <- "#1F2937"
C_MUTED <- "#6B7280"; C_LINE <- "#E5E7EB"; C_HEAD <- "#F3F6FA"; C_GROUP <- "#CBD5E1"
FONT <- "font-family:'Segoe UI',Helvetica,Arial,sans-serif;"

e <- function(x) as.character(esc_html(coalesce(as.character(x), "")))

# Small rounded label. Outlook desktop shows square corners; everything else rounds them.
PILL_TONES <- list(
  green = c("#DCFCE7", "#166534"), blue  = c("#DBEAFE", "#1E40AF"),
  amber = c("#FEF3C7", "#92400E"), red   = c("#FEE2E2", "#991B1B"),
  grey  = c("#F1F5F9", "#475569")
)
pill <- function(txt, tone = "grey") {
  tone <- rep_len(tone, length(txt))
  tone[is.na(tone) | !tone %in% names(PILL_TONES)] <- "grey"
  bg <- vapply(tone, function(t) PILL_TONES[[t]][1], character(1))
  fg <- vapply(tone, function(t) PILL_TONES[[t]][2], character(1))
  sprintf('<span style="display:inline-block;padding:2px 8px;border-radius:999px;background:%s;color:%s;font-size:11px;font-weight:600;white-space:nowrap;">%s</span>',
          bg, fg, e(txt))
}
muted  <- function(txt) sprintf('<span style="color:%s;">%s</span>', C_MUTED, e(txt))
strong <- function(txt) sprintf('<span style="font-weight:600;">%s</span>', e(txt))

# "Calgary, AB" / "Ashburn, VA" / "Navi Mumbai, India" (province or state when the sheet has one)
loc_txt <- function(market, state, country) {
  st <- str_squish(coalesce(as.character(state), ""))
  where <- txt_ifelse(nzchar(st), st, coalesce(as.character(country), ""))
  txt_ifelse(nzchar(where), paste0(coalesce(market, ""), ", ", where), coalesce(market, ""))
}

mw_txt <- function(x) txt_ifelse(is.na(x), "\u2013", paste0(fmt_mw(x), " MW"))
q_txt  <- function(d) txt_ifelse(is.na(d), "\u2013", quarter_label(d))

# "36 → 48 MW  +12"  (old value muted, new value bold, change as a coloured pill)
mw_change_html <- function(before, after) {
  d <- after - before
  delta <- txt_ifelse(is.na(d), "",
                      paste0("&nbsp;", pill(paste0(txt_ifelse(d > 0, "+", "\u2212"), fmt_mw(abs(d))),
                                            txt_ifelse(d > 0, "green", "red"))))
  paste0(muted(txt_ifelse(is.na(before), "none", fmt_mw(before))), " \u2192 ",
         strong(txt_ifelse(is.na(after), "none", mw_txt(after))), delta)
}

# "Q1 2027 · 20 MW", or a muted "Not listed"
first_avail_html <- function(d, mw) {
  txt_ifelse(is.na(d), muted("Not listed"),
             paste0(strong(q_txt(d)), muted(paste0(" \u00b7 ", mw_txt(mw)))))
}

quarters_between <- function(a, b) {
  as.integer(round((as.integer(format(b, "%Y")) * 12 + as.integer(format(b, "%m")) -
                    as.integer(format(a, "%Y")) * 12 - as.integer(format(a, "%m"))) / 3))
}
plural <- function(n, word) paste0(n, " ", word, txt_ifelse(n == 1, "", "s"))

# MW available in each site's first available quarter
first_mw <- now$pipe %>%
  group_by(key) %>%
  slice_min(Quarter_Date, n = 1, with_ties = FALSE) %>%
  ungroup() %>%
  select(key, First_MW = MW)

cq_start   <- current_quarter_start(TODAY)
window_end <- seq(cq_start, by = "3 months", length.out = COMING_SOON_QUARTERS + 1)[COMING_SOON_QUARTERS + 1]

big_change <- function(a, b) {
  (is.na(a) != is.na(b)) | (!is.na(a) & !is.na(b) & abs(a - b) >= MIN_MW_CHANGE)
}

sec <- list()

if (!is.null(baseline)) {
  ns <- now$sites
  bs <- baseline$data$sites

  # Row counts per operator + market, now vs. last week
  counts <- full_join(
    ns %>% select(key, Rows_now = Rows),
    bs %>% select(key, Rows_before = Rows),
    by = "key"
  ) %>%
    mutate(Rows_now = coalesce(Rows_now, 0L), Rows_before = coalesce(Rows_before, 0L),
           Added = pmax(Rows_now - Rows_before, 0L),
           Dropped = pmax(Rows_before - Rows_now, 0L))

  n_added   <- sum(counts$Added)
  n_dropped <- sum(counts$Dropped)

  # 1. New sites: every added row counts, including extra suites/floors
  #    at an operator + market that was already in the sheet.
  #    Biggest operators first, so a 15-site expansion reads as one block.
  sec$new <- ns %>%
    inner_join(counts %>% filter(Added > 0), by = "key") %>%
    left_join(first_mw, by = "key") %>%
    group_by(Operator) %>% mutate(op_added = sum(Added)) %>% ungroup() %>%
    arrange(desc(op_added), Operator, desc(Added), desc(coalesce(Total_MW, 0))) %>%
    transmute(
      Operator = strong(Operator),
      Location = e(loc_txt(Market, State, Country)),
      Sites = txt_ifelse(Rows_before == 0,
                         paste0(strong(Added), "&nbsp; ", pill("New market", "green")),
                         paste0(strong(paste0("+", Added)), "&nbsp; ", muted(paste0("now ", Rows_now)))),
      Capacity = strong(mw_txt(Total_MW)),
      `First availability` = first_avail_html(First_Q_Date, First_MW)
    )

  both <- ns %>%
    inner_join(
      bs %>% select(key, Total_before = Total_MW, FirstQ_before = First_Q_Date, Peak_before = Peak_MW),
      by = "key"
    ) %>%
    left_join(counts %>% select(key, Rows_before), by = "key") %>%
    mutate(sold_out = !is.na(Peak_before) & is.na(Peak_MW),
           same_rows = Rows == Rows_before)

  # 2a. Capacity movers (edits to existing rows; added/removed rows are
  #     already listed in New sites / Removed)
  sec$movers <- both %>%
    filter(!sold_out, same_rows, big_change(Total_MW, Total_before) | big_change(Peak_MW, Peak_before)) %>%
    mutate(size = pmax(abs(coalesce(Total_MW, 0) - coalesce(Total_before, 0)),
                       abs(coalesce(Peak_MW, 0) - coalesce(Peak_before, 0)))) %>%
    arrange(desc(size)) %>%
    transmute(
      Operator = strong(Operator),
      Location = e(loc_txt(Market, State, Country)),
      `Total capacity` = txt_ifelse(big_change(Total_MW, Total_before),
                                    mw_change_html(Total_before, Total_MW), muted(mw_txt(Total_MW))),
      Upcoming = txt_ifelse(big_change(Peak_MW, Peak_before),
                            mw_change_html(Peak_before, Peak_MW), muted(mw_txt(Peak_MW)))
    )

  # 2b. Timeline changes (only when last week's date was still in the future,
  #     so clearing out a quarter that already passed doesn't count as a slip)
  sec$timeline <- both %>%
    filter(!is.na(First_Q_Date), !is.na(FirstQ_before),
           First_Q_Date != FirstQ_before, FirstQ_before >= cq_start) %>%
    mutate(slipped = First_Q_Date > FirstQ_before,
           n_q = abs(quarters_between(FirstQ_before, First_Q_Date))) %>%
    arrange(desc(slipped), desc(n_q), FirstQ_before) %>%
    transmute(
      Operator = strong(Operator),
      Location = e(loc_txt(Market, State, Country)),
      Was = muted(q_txt(FirstQ_before)),
      Now = strong(q_txt(First_Q_Date)),
      Change = txt_ifelse(slipped,
                          pill(paste("Slipped", plural(n_q, "qtr")), "amber"),
                          pill(paste(plural(n_q, "qtr"), "earlier"), "green"))
    )

  # 3. Removed or sold out
  removed <- bs %>%
    inner_join(counts %>% filter(Dropped > 0), by = "key") %>%
    arrange(desc(Dropped)) %>%
    transmute(Operator = strong(Operator),
              Location = e(loc_txt(Market, State, Country)),
              `Last week` = e(mw_txt(Total_MW)),
              Status = txt_ifelse(Rows_now == 0,
                                  pill(paste(plural(Dropped, "site"), "removed"), "red"),
                                  paste0(pill(paste(Dropped, "removed"), "grey"), "&nbsp; ",
                                         muted(paste(Rows_now, "left")))))
  sold <- both %>%
    filter(sold_out) %>%
    transmute(Operator = strong(Operator),
              Location = e(loc_txt(Market, State, Country)),
              `Last week` = e(paste0(mw_txt(Peak_before), " upcoming")),
              Status = pill("Sold out", "amber"))
  # Both tables are display-only text; forcing text means an empty table
  # can never clash with a filled one
  both_sold_n <- nrow(sold)
  sec$gone <- bind_rows(
    sold %>% mutate(across(everything(), as.character)),
    removed %>% mutate(across(everything(), as.character))
  )
}

# 4. Coming online soon (first available quarter falls in the window)
mw_by_window <- now$pipe %>%
  filter(Quarter_Date <= window_end) %>%
  group_by(key) %>%
  summarise(MW_window = max(MW), .groups = "drop")

coming_raw <- now$sites %>%
  filter(!is.na(First_Q_Date), First_Q_Date >= cq_start, First_Q_Date <= window_end) %>%
  left_join(mw_by_window, by = "key") %>%
  arrange(First_Q_Date, desc(MW_window))

sec$coming <- coming_raw %>%
  transmute(
    Operator = strong(Operator),
    Location = e(loc_txt(Market, State, Country)),
    `Available` = txt_ifelse(First_Q_Date == cq_start, paste0(strong(q_txt(First_Q_Date)), "&nbsp; ", pill("This quarter", "blue")),
                                  strong(q_txt(First_Q_Date))),
    !!paste0("MW by ", quarter_label(window_end)) := strong(mw_txt(MW_window))
  )

# 5. Stale rows
stale_raw <- now$sites %>%
  left_join(site_hist %>% select(key, first_seen, last_changed), by = "key") %>%
  left_join(upload_dates, by = "key") %>%
  mutate(
    tracked = as.Date(last_changed),
    Last_Update = pmax(Upload_Date, tracked, na.rm = TRUE),
    known = !is.na(Last_Update),
    Last_Update = coalesce(Last_Update, as.Date(first_seen)),   # unknown -> since tracking began
    Days = as.integer(TODAY - Last_Update)
  ) %>%
  filter(Days >= STALE_DAYS) %>%
  arrange(desc(Days))

sec$stale <- stale_raw %>%
  transmute(
    Operator = strong(Operator),
    Location = e(loc_txt(Market, State, Country)),
    `Last updated` = txt_ifelse(known, e(format(Last_Update, "%b %d, %Y")),
                                muted(paste0("before ", format(Last_Update, "%b %d, %Y")))),
    Days = txt_ifelse(Days >= 90, pill(paste(Days, "days"), "amber"), strong(paste(Days, "days")))
  )

tracking_started <- suppressWarnings(min(as.Date(site_hist$first_seen), na.rm = TRUE))

# ============================================================
# HTML EMAIL (table-based + inline styles so Outlook renders it)
# ============================================================

# Cells arrive as ready-made HTML (escaped by e()/pill()/strong()/muted()).
# Columns named in `right` are right-aligned (numbers line up).
# Rows for the same operator are kept together and the name is shown once.
html_table <- function(df, empty_msg, right = character()) {
  if (is.null(df) || nrow(df) == 0) {
    return(sprintf('<p style="%smargin:4px 0 0;padding:12px 14px;font-size:13px;color:%s;background:%s;border-radius:6px;">%s</p>',
                   FONT, C_MUTED, C_HEAD, esc_html(empty_msg)))
  }

  if (names(df)[1] == "Operator") {
    df <- df[order(match(df$Operator, unique(df$Operator))), , drop = FALSE]   # stable grouping
  }
  extra <- nrow(df) - MAX_ROWS
  df <- head(df, MAX_ROWS)
  is_right <- names(df) %in% right

  th <- paste0(sprintf(
    '<th align="%s" style="%spadding:8px 12px;font-size:10px;font-weight:700;color:%s;text-transform:uppercase;letter-spacing:.6px;background:%s;border-bottom:1px solid %s;white-space:nowrap;">%s</th>',
    ifelse(is_right, "right", "left"), FONT, C_MUTED, C_HEAD, C_LINE, esc_html(names(df))), collapse = "")

  ops <- if (names(df)[1] == "Operator") df$Operator else rep(NA_character_, nrow(df))

  rows <- vapply(seq_len(nrow(df)), function(i) {
    new_group <- i == 1 || !identical(ops[i], ops[i - 1])
    top <- if (i > 1 && new_group && !is.na(ops[i])) sprintf("border-top:1px solid %s;", C_GROUP) else ""
    cells <- vapply(seq_along(df), function(j) {
      v <- as.character(df[[j]][i])
      if (j == 1 && !is.na(ops[i]) && !new_group) v <- ""          # operator shown once per group
      sprintf('<td align="%s" valign="top" style="%spadding:9px 12px;font-size:13px;line-height:1.4;color:%s;border-bottom:1px solid %s;%s">%s</td>',
              if (is_right[j]) "right" else "left", FONT, C_TEXT, C_LINE, top,
              if (is.na(v) || !nzchar(v)) "&nbsp;" else v)
    }, character(1))
    paste0("<tr>", paste(cells, collapse = ""), "</tr>")
  }, character(1))

  more <- if (extra > 0) {
    sprintf('<p style="%smargin:8px 0 0;font-size:12px;color:%s;">+ %d more %s</p>',
            FONT, C_MUTED, extra,
            if (nzchar(DASHBOARD_URL)) sprintf('&mdash; <a href="%s" style="color:%s;font-weight:600;text-decoration:none;">see all in the dashboard &rarr;</a>', DASHBOARD_URL, C_BLUE) else "in the dashboard")
  } else ""

  paste0('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid ',
         C_LINE, ';">', "<tr>", th, "</tr>", paste(rows, collapse = ""), "</table>", more)
}

section_html <- function(title, hint, body) {
  sprintf(paste0(
    '<tr><td style="padding:22px 28px 0;">',
    '<p style="%smargin:0;font-size:16px;font-weight:700;color:%s;">%s</p>',
    '<p style="%smargin:2px 0 10px;font-size:12px;color:%s;">%s</p>%s</td></tr>'),
    FONT, C_NAVY, esc_html(title), FONT, C_MUTED, esc_html(hint), body)
}

kpi_cell <- function(value, label) {
  sprintf(paste0(
    '<td width="25%%" align="center" style="padding:14px 6px;background:#FFFFFF;border:1px solid %s;">',
    '<p style="%smargin:0;font-size:22px;font-weight:700;color:%s;">%s</p>',
    '<p style="%smargin:2px 0 0;font-size:11px;color:%s;">%s</p></td>'),
    C_LINE, FONT, C_BLUE, esc_html(value), FONT, C_MUTED, esc_html(label))
}

n_or_dash <- function(df) if (is.null(df)) "–" else as.character(nrow(df))
coming_mw <- sum(coming_raw$MW_window, na.rm = TRUE)

kpis <- paste0(
  kpi_cell(if (is.null(baseline)) "–" else format(n_added, big.mark = ","), "New sites"),
  kpi_cell(if (is.null(baseline)) "–" else as.character(nrow(sec$movers) + nrow(sec$timeline)), "Capacity / timeline changes"),
  kpi_cell(paste0(fmt_mw(coming_mw), " MW"), paste0("Coming online by ", quarter_label(window_end))),
  kpi_cell(if (is.null(baseline)) "–" else format(n_dropped + sum(both_sold_n), big.mark = ","), "Removed / sold out")
)

since_txt <- if (is.null(baseline)) {
  "No earlier snapshot found, so changes will start showing next week."
} else {
  sprintf("Changes since %s (%s).", format(baseline$as_of, "%b %d, %Y"), baseline$source)
}
renamed_table <- rename_summary %>%
  arrange(desc(n_sites)) %>%
  transmute(`Old name` = muted(old_op), `New name` = strong(new_op), Sites = strong(n_sites))

no_base <- "No earlier snapshot to compare with yet."

body_sections <- paste0(
  section_html("New sites", "Added to the inventory since last week.",
               if (is.null(baseline)) html_table(NULL, no_base) else html_table(sec$new, "No new sites this week.", right = "Capacity")),
  # Only appears in weeks with a rename
  if (nrow(renamed_table) > 0) {
    section_html("Renamed operators",
                 "Same sites under a new operator name. Their history carries over, so they aren't counted as new or removed.",
                 html_table(renamed_table, "", right = "Sites"))
  } else "",
  section_html("Capacity changes", "Total or upcoming capacity went up or down.",
               if (is.null(baseline)) html_table(NULL, no_base) else html_table(sec$movers, "No capacity changes this week.")),
  section_html("Timeline changes", "The first quarter capacity is available moved. Slips are worth a heads-up to affected clients.",
               if (is.null(baseline)) html_table(NULL, no_base) else html_table(sec$timeline, "No timeline changes this week.")),
  section_html("Removed or sold out", "Stop pitching these, or check in with the operator.",
               if (is.null(baseline)) html_table(NULL, no_base) else html_table(sec$gone, "Nothing removed or sold out this week.", right = "Last week")),
  section_html(paste0("Coming online soon (", quarter_label(cq_start), " – ", quarter_label(window_end), ")"),
               "Sites whose capacity first becomes available in this window, earliest first.",
               html_table(sec$coming, "No sites first coming online in this window.", right = names(sec$coming)[4])),
  section_html(paste0("Not updated in ", STALE_DAYS, "+ days"),
               "Worth a check-in with the operator to confirm the numbers.",
               html_table(sec$stale,
                          if (is.finite(tracking_started) && as.integer(TODAY - tracking_started) < STALE_DAYS &&
                              nrow(upload_dates) == 0) {
                            sprintf("Tracking started %s. Fill in \"Capacity Upload date\" in the sheet to see stale rows sooner.",
                                    format(tracking_started, "%b %d, %Y"))
                          } else "Everything has been updated recently.",
                          right = "Days"))
)

links <- c(
  if (nzchar(DASHBOARD_URL)) sprintf('<a href="%s" style="color:%s;font-weight:600;">Open the dashboard</a>', DASHBOARD_URL, C_BLUE),
  if (nzchar(REPORT_SHAREPOINT_URL)) sprintf('<a href="%s" style="color:%s;font-weight:600;">Master Excel on SharePoint</a>', REPORT_SHAREPOINT_URL, C_BLUE)
)
links_html <- if (length(links) > 0) {
  sprintf('<tr><td style="padding:22px 28px 0;%sfont-size:13px;">%s</td></tr>', FONT, paste(links, collapse = " &nbsp;&middot;&nbsp; "))
} else ""

test_banner <- if (TEST_MODE) {
  sprintf('<tr><td style="padding:8px 28px;background:#FEF3C7;%sfont-size:12px;color:#92400E;font-weight:600;">TEST &mdash; sent only to you. Set TEST_MODE &lt;- FALSE to send to the sales team.</td></tr>', FONT)
} else ""

html <- paste0(
  '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>',
  '<body style="margin:0;padding:0;background:#EEF1F5;">',
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF1F5;"><tr><td align="center" style="padding:20px 10px;">',
  '<table role="presentation" width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;background:#FFFFFF;border-collapse:collapse;">',
  test_banner,
  sprintf(paste0('<tr><td style="padding:22px 28px;background:%s;">',
                 '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0"><tr>{{LOGO_CELL}}<td valign="middle">',
                 '<p style="%smargin:0;font-size:12px;color:#9DB4D3;letter-spacing:.6px;text-transform:uppercase;">{{EYEBROW}}</p>',
                 '<p style="%smargin:4px 0 0;font-size:22px;font-weight:700;color:#FFFFFF;">Weekly Capacity Report</p>',
                 '<p style="%smargin:4px 0 0;font-size:13px;color:#C9D6E8;">%s &middot; %s sites tracked%s</p>',
                 '</td></tr></table></td></tr>'),
          C_NAVY, FONT, FONT, FONT, format(TODAY, "%B %d, %Y"), format(sum(now$sites$Rows), big.mark = ","),
          if (is.null(baseline)) "" else {
            net <- sum(now$sites$Rows) - sum(baseline$data$sites$Rows)
            sprintf(" (%s%s since %s)", if (net >= 0) "+" else "", format(net, big.mark = ","), format(baseline$as_of, "%b %d"))
          }),
  sprintf('<tr><td style="padding:18px 28px 0;"><table role="presentation" width="100%%" cellpadding="0" cellspacing="6" style="background:%s;"><tr>%s</tr></table>', C_HEAD, kpis),
  sprintf('<p style="%smargin:8px 0 0;font-size:12px;color:%s;">%s</p></td></tr>', FONT, C_MUTED, esc_html(since_txt)),
  body_sections,
  links_html,
  sprintf('<tr><td style="padding:26px 28px 24px;%sfont-size:11px;color:%s;">Generated automatically from the DC Capacity Dashboard data. Questions or corrections: reply to this email.</td></tr>', FONT, C_MUTED),
  '</table></td></tr></table></body></html>'
)

# ---------- Logo ----------
# The preview embeds the image directly. The email attaches it as an inline
# image (cid:), because Outlook desktop and Gmail block embedded data: images.

png_size <- function(path) {
  b <- readBin(path, "raw", 24)
  png_sig <- as.raw(c(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
  if (length(b) < 24 || !identical(b[1:8], png_sig)) return(NULL)
  to_int <- function(x) sum(as.numeric(as.integer(x)) * 256^(3:0))
  c(w = to_int(b[17:20]), h = to_int(b[21:24]))
}

logo <- NULL
if (nzchar(LOGO_FILE) && file.exists(LOGO_FILE)) {
  ext  <- tolower(tools::file_ext(LOGO_FILE))
  mime <- switch(ext, png = "image/png", jpg = , jpeg = "image/jpeg", gif = "image/gif", NA_character_)
  if (is.na(mime)) {
    cat("Logo skipped: use a .png, .jpg or .gif file.\n")
  } else {
    bytes <- readBin(LOGO_FILE, "raw", file.info(LOGO_FILE)$size)
    size  <- if (ext == "png") png_size(LOGO_FILE) else NULL
    logo <- list(
      name   = basename(LOGO_FILE),
      mime   = mime,
      b64    = gsub("[\r\n]", "", jsonlite::base64_enc(bytes)),
      height = LOGO_HEIGHT,
      # Outlook needs an explicit width, so keep the logo's proportions
      width  = if (is.null(size)) NA else round(LOGO_HEIGHT * size[["w"]] / size[["h"]])
    )
  }
} else if (nzchar(LOGO_FILE)) {
  cat("Logo not found at", LOGO_FILE, "- sending without it.\n")
}

with_logo <- function(html, src) {
  if (is.null(logo) || is.null(src)) {
    cell <- ""
    eyebrow <- "ByteBridge &middot; Data Center Capacity"
  } else {
    size_attr <- if (is.na(logo$width)) sprintf('height="%d"', logo$height) else
      sprintf('height="%d" width="%d"', logo$height, logo$width)
    size_css <- if (is.na(logo$width)) sprintf("height:%dpx;", logo$height) else
      sprintf("height:%dpx;width:%dpx;", logo$height, logo$width)
    cell <- sprintf(paste0(
      '<td valign="middle" style="padding-right:18px;%s">',
      '<img src="%s" %s alt="ByteBridge" style="display:block;border:0;outline:none;%sborder-radius:6px;"></td>'),
      if (is.na(logo$width)) "" else sprintf("width:%dpx;", logo$width), src, size_attr, size_css)
    eyebrow <- "Data Center Capacity"     # the logo already says ByteBridge
  }
  html <- sub("{{LOGO_CELL}}", cell, html, fixed = TRUE)
  sub("{{EYEBROW}}", eyebrow, html, fixed = TRUE)
}

html_template <- html
html_preview  <- with_logo(html_template, if (!is.null(logo)) paste0("data:", logo$mime, ";base64,", logo$b64))
html_email    <- with_logo(html_template, if (!is.null(logo)) "cid:bytebridge_logo")
html_no_logo  <- with_logo(html_template, NULL)

preview_file <- file.path(getwd(), "weekly_report_preview.html")
writeLines(html_preview, preview_file, useBytes = TRUE)
cat("Preview saved:", preview_file, "\n")
# Open the preview when run by hand in RStudio. The opener is looked up by
# name so the shinyapps.io publish check doesn't flag this script (it only
# ever runs on your Mac, never on the published site).
if (interactive()) {
  open_in_browser <- get("browseURL", envir = asNamespace("utils"))
  open_in_browser(preview_file)
}

# ============================================================
# SEND
# ============================================================

subject <- paste0(if (TEST_MODE) "[TEST] " else "",
                  "Weekly DC Capacity Report – ", format(TODAY, "%b %d, %Y"))
recipients <- if (TEST_MODE) MY_EMAIL else SALES_TEAM

sent <- FALSE
if (SEND_EMAIL) {
  if (length(recipients) == 0) stop("SALES_TEAM is empty - add the 4 email addresses at the top of the script.")
  if (!requireNamespace("Microsoft365R", quietly = TRUE)) {
    stop("Microsoft365R is not installed. Run: install.packages(\"Microsoft365R\")")
  }

  sent <- tryCatch({
    outl <- Microsoft365R::get_business_outlook()
    em <- outl$create_email(body = if (is.null(logo)) html_no_logo else html_email,
                            content_type = "html",
                            subject = subject, to = recipients,
                            cc = if (!TEST_MODE) MY_EMAIL else NULL)

    # Attach the logo as an inline image the HTML points to (cid:bytebridge_logo)
    if (!is.null(logo)) {
      ok <- tryCatch({
        em$do_operation("attachments", http_verb = "POST", body = list(
          `@odata.type` = "#microsoft.graph.fileAttachment",
          name          = logo$name,
          contentType   = logo$mime,
          contentBytes  = logo$b64,
          contentId     = "bytebridge_logo",
          isInline      = TRUE
        ))
        TRUE
      }, error = function(e) {
        cat("Could not attach the logo (", conditionMessage(e), ") - sending without it.\n")
        FALSE
      })
      if (!ok) em$set_body(html_no_logo, content_type = "html")
    }

    em$send()
    cat("Sent to:", paste(recipients, collapse = ", "), "\n")
    TRUE
  }, error = function(e) {
    cat("\nSENDING FAILED:", conditionMessage(e), "\n")
    cat("If the error mentions admin consent / approval, your IT admin needs to allow\n",
        "the Microsoft365R app in Microsoft 365 (or we can switch to another send method).\n",
        "The preview file above was still created.\n")
    FALSE
  })
}

# Only a real send moves the baseline forward, so test runs can be repeated
if (sent && !TEST_MODE) {
  save_own_baseline(now)
  cat("Saved this week's data as next week's baseline.\n")
}
