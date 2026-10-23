# ============================================================
# Data Center Capacity Dashboard
#
# MASTER DATA:   DC_Inventory_Master_COMBINED_THANOS (9).xlsx
#                on SharePoint (DataCenterCapacity > Documents >
#                Glen and Jacob), synced to this Mac by OneDrive
# APP COPY:      DC.xlsx  - a symlink in this folder that points to
#                the synced master file (not a separate copy)
# RUNTIME DATA:  dc_data.sqlite
#
# Workflow:
# 1. Anyone edits the master workbook on SharePoint
# 2. OneDrive syncs it to this Mac automatically (no download/rename)
# 3. Launch the app - it imports DC.xlsx automatically at startup
#    (only if the file changed since the last import)
# 4. (Optional) "Refresh from DC.xlsx" in Version History if the
#    master sheet changes while the app is already running
# 5. Republish the app so the hosted version gets the new data
#
# If the master file is ever renamed on SharePoint, recreate the link
# in Terminal:
#   rm -f ~/Desktop/shiny/DC.xlsx
#   ln -s "<full path to the new synced file>" ~/Desktop/shiny/DC.xlsx
#
# SQLite is the fast runtime database. Geocoding happens during
# import and reuses cached coordinates.
#
# Maps: every map has a Map / Satellite switch (top-right corner).
# Every map always opens on "Map" (CARTO Positron); Satellite uses
# Esri World Imagery with place-name labels on top.
#
# Data rules worth knowing:
#   * Quarter columns are detected from the sheet headers
#     (any column named like "Q3 2026"), so adding "Q1 2030"
#     just works and quarters always sort by date.
#   * Capacity text is read unit-aware: "1.2 GW" = 1,200 MW,
#     "500 kW" = 0.5 MW, "Phase 1: 18 MW, Phase 2: 24 MW" = 42 MW,
#     "20-40 MW" = 30 MW estimate with a 20-40 MW range.
#   * Anything that cannot be read cleanly is listed under
#     Version History > Data quality check.
# ============================================================

# ---------- Packages ----------

library(shiny)
library(bslib)
library(leaflet)
library(DT)
library(plotly)
library(dplyr)
library(stringr)
library(DBI)
library(RSQLite)
library(readxl)
library(tidyr)
library(jsonlite)

# ---------- File locations & settings ----------

DB_PATH <- "dc_data.sqlite"
MASTER_FILE <- "DC.xlsx"

# Optional: link to the master workbook on SharePoint (shown in the header).
# Leave as "" to hide the link.
SHAREPOINT_URL <- ""

# Bump this whenever the import/parsing logic changes. It is part of the
# import signature, so the next launch re-imports DC.xlsx even if the file
# itself has not changed.
IMPORT_SCHEMA_VERSION <- "3"

## carto map api key
CARTO_KEY <- Sys.getenv("CARTO_API_KEY")
if (!nzchar(CARTO_KEY) && file.exists("carto_key.txt")) {
  CARTO_KEY <- trimws(readLines("carto_key.txt", n = 1, warn = FALSE))
}
CARTO_POSITRON_URL <- paste0(
  "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
  if (nzchar(CARTO_KEY)) paste0("?key=", CARTO_KEY) else ""
)
CARTO_ATTRIBUTION <- paste0(
  '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> ',
  'contributors &copy; <a href="https://carto.com/attributions">CARTO</a>'
)

## Satellite view: place-name labels drawn on top of the imagery.
## Uses the same Carto key as the regular map (carto_key.txt).
CARTO_LABELS_URL <- paste0(
  "https://{s}.basemaps.cartocdn.com/dark_only_labels/{z}/{x}/{y}{r}.png",
  if (nzchar(CARTO_KEY)) paste0("?key=", CARTO_KEY) else ""
)

# ---------- Basemap switch ----------
# The first name is the default every map opens on.
BASEMAP_DEFAULT   <- "Map"
BASEMAP_SATELLITE <- "Satellite"

# Password required for "Refresh from DC.xlsx" and "Restore version".
REFRESH_PASSWORD <- Sys.getenv("DC_REFRESH_PASSWORD", unset = "bytebt")

# ---------- Quarter columns ----------
# Any sheet column whose header looks like "Q3 2026" is an upcoming-capacity
# column. There is no fixed list: the sheet decides.

QUARTER_PATTERN <- "^Q[1-4] \\d{4}$"

# ---------- Optional master-file columns ----------

MASTER_OPTIONAL_COLS <- c(
  "Market", "Address", "City", "State/Province", "Region", "Capacity Upload date",
  "Utility Rate ($/kWh)", "Price ($/kW)", "Cooling", "PUE",
  "Tax Incentives", "Deal Reg", "Notes", "Contacts"
)

# ---------- Country centroids (used only when a row has no city) ----------

country_centroids <- tribble(
  ~Country,               ~Latitude, ~Longitude,
  "United States",          39.5,      -98.5,
  "Canada",                 56.1,     -106.3,
  "United Kingdom",         54.0,       -2.0,
  "Ireland",                53.4,       -8.2,
  "Germany",                51.2,       10.4,
  "France",                 46.6,        2.2,
  "Netherlands",            52.1,        5.3,
  "Spain",                  40.0,       -4.0,
  "Italy",                  42.8,       12.6,
  "Sweden",                 62.0,       15.0,
  "Norway",                 60.5,        8.5,
  "Poland",                 52.0,       19.0,
  "Switzerland",            46.8,        8.2,
  "Portugal",               39.6,       -8.0,
  "Belgium",                50.6,        4.7,
  "Denmark",                56.0,       10.0,
  "Finland",                64.0,       26.0,
  "Singapore",               1.35,     103.8,
  "Japan",                  36.2,      138.3,
  "India",                  22.0,       79.0,
  "Australia",             -25.3,      133.8,
  "Brazil",                -14.2,      -51.9,
  "Mexico",                 23.6,     -102.5,
  "United Arab Emirates",   24.0,       54.0,
  "South Korea",            36.5,      127.8,
  "China",                  35.9,      104.2,
  "Hong Kong",              22.3,      114.2,
  "Vietnam",                14.1,      108.3,
  "Indonesia",              -0.8,      113.9,
  "Thailand",               15.9,      101.0,
  "Malaysia",                4.2,      101.9,
  "Saudi Arabia",           24.0,       45.0,
  "South Africa",          -30.6,       22.9,
  "Uruguay",               -32.5,      -55.8,
  "Sweden/Norway",          61.3,       11.8
)

MANUAL_COORDS <- tribble(
  ~key,                           ~Latitude, ~Longitude,
  "Santa Clara|CA|United States",  37.3541,  -121.9552
)

# ============================================================
# SMALL GENERAL HELPERS
# ============================================================

`%||%` <- function(a, b) if (is.null(a)) b else a

esc_html <- function(x) htmltools::htmlEscape(enc2utf8(as.character(x)))

# ============================================================
# QUARTER HELPERS
# ============================================================

# "Q2 2027" -> 2027-04-01
quarter_to_date <- function(q) {
  m <- str_match(q, "^Q([1-4])\\s+(\\d{4})$")
  qn <- as.integer(m[, 2])
  yr <- as.integer(m[, 3])
  month <- (qn - 1L) * 3L + 1L
  
  as.Date(ifelse(
    is.na(qn) | is.na(yr),
    NA_character_,
    sprintf("%04d-%02d-01", yr, month)
  ))
}

# 2027-04-01 -> "Q2 2027"
quarter_label <- function(d) {
  m <- as.integer(format(d, "%m"))
  ifelse(is.na(d), NA_character_, paste0("Q", (m - 1L) %/% 3L + 1L, " ", format(d, "%Y")))
}

# First day of the quarter that contains `d` (today by default)
current_quarter_start <- function(d = Sys.Date()) {
  m <- as.integer(format(d, "%m"))
  y <- as.integer(format(d, "%Y"))
  as.Date(sprintf("%04d-%02d-01", y, (m - 1L) %/% 3L * 3L + 1L))
}

# ============================================================
# MW FORMATTING HELPERS (estimate + min-max range text)
# ============================================================

fmt_mw <- function(x) {
  ifelse(is.na(x), "\u2013", prettyNum(round(x, 1), big.mark = ","))
}

# "20\u201340" when there is a real range, "" when min == max or unknown
mw_span_txt <- function(lo, hi) {
  ifelse(
    is.na(lo) | is.na(hi) | abs(hi - lo) < 0.05,
    "",
    paste0(fmt_mw(lo), "\u2013", fmt_mw(hi))
  )
}

# "30 MW (range 20\u201340 MW)" or "18 MW"
mw_range_txt <- function(est, lo, hi) {
  span <- mw_span_txt(lo, hi)
  as.character(ifelse(
    is.na(est),
    "No data",
    paste0(fmt_mw(est), " MW", ifelse(nzchar(span), paste0(" (range ", span, " MW)"), ""))
  ))
}

# ============================================================
# CAPACITY PARSING
# ============================================================
# Reads free text such as:
#   "20-40 MW"                          -> min 20, max 40, est 30
#   "3 MW + 6 MW"                       -> 9
#   "Phase 1: 18 MW, Phase 2: 24 MW"    -> 42   (phase numbers are not MW)
#   "1.2 GW"                            -> 1,200
#   "500 kW"                            -> 0.5
#   "1,200 MW"                          -> 1,200
#   "20-40"  (no unit)                  -> 30   (the column is "MW")
# Every value is returned as min / max / estimate (midpoint) plus a
# plain-English flag when the text could not be read cleanly.

CAP_UNIT <- "(?:gigawatts?|megawatts?|kilowatts?|mwe|mva|gw|mw|kw)"
CAP_NUM  <- "\\d+(?:\\.\\d+)?"

# number [unit] [ - | to  number [unit] ]
CAP_TOKEN_RE <- regex(
  paste0(
    "(", CAP_NUM, ")\\s*(", CAP_UNIT, "\\b)?",
    "(?:\\s*(?:-|\u2013|\u2014|\\bto\\b)\\s*(", CAP_NUM, ")\\s*(", CAP_UNIT, "\\b)?)?"
  ),
  ignore_case = TRUE
)

# "Phase 2", "Building 3", "Q3" ... labels whose numbers are not capacity
CAP_LABEL_RE <- regex(
  paste0(
    "\\b(?:phase|stage|building|bldg|block|hall|module|tranche|campus|site|unit|pod|suite|floor|level|q)",
    "\\s*#?\\s*\\d+(?:\\s*(?:&|/|and)\\s*\\d+|-\\d+)?"
  ),
  ignore_case = TRUE
)

CAP_NOISE_RE <- regex(
  "\\b(?:approx(?:imately)?|est(?:imated)?|about|circa|around|roughly)\\b\\.?",
  ignore_case = TRUE
)

# Wording that makes the number unreliable on its own
CAP_AMBIG_STRONG_RE <- regex(
  "\\b(?:or|up to|at least|at most|min|max|minimum|maximum)\\b|\\d\\s*[x\u00d7]\\s*\\d",
  ignore_case = TRUE
)
# Wording that may mean numbers overlap (only matters with 2+ numbers)
CAP_AMBIG_MULTI_RE <- regex(
  "\\b(?:incl\\w*|of which|total|subtotal)\\b|\\([^)]*\\d",
  ignore_case = TRUE
)

cap_unit_factor <- function(u) {
  u <- str_to_lower(u)
  case_when(
    is.na(u) ~ 1,
    str_starts(u, "g") ~ 1000,
    str_starts(u, "k") ~ 0.001,
    TRUE ~ 1
  )
}

parse_capacity_one <- function(val) {
  none <- list(min = NA_real_, max = NA_real_, est = NA_real_, flag = NA_character_)
  
  if (length(val) != 1 || is.na(val)) return(none)
  s <- str_squish(as.character(val))
  if (!nzchar(s)) return(none)
  
  # 1,200 -> 1200 (thousands separator only), then drop labels and filler words
  s <- str_replace_all(s, "(?<=\\d),(?=\\d{3}(?!\\d))", "")
  s <- str_replace_all(s, CAP_LABEL_RE, " ")
  s <- str_replace_all(s, CAP_NOISE_RE, " ")
  s <- str_squish(s)
  
  no_number <- function() {
    r <- none
    r$flag <- "No number found"
    r
  }
  
  m <- str_match_all(s, CAP_TOKEN_RE)[[1]]
  if (nrow(m) == 0) return(no_number())
  
  blank_to_na <- function(x) {
    x[!is.na(x) & !nzchar(x)] <- NA_character_
    x
  }
  
  lo_v <- as.numeric(m[, 2])
  hi_v <- suppressWarnings(as.numeric(m[, 4]))
  hi_v <- ifelse(is.na(hi_v), lo_v, hi_v)
  u1 <- blank_to_na(m[, 3])
  u2 <- blank_to_na(m[, 5])
  
  has_unit <- !is.na(u1) | !is.na(u2)
  
  # A bare 4-digit number between 1990 and 2100 is a year, not capacity
  is_year <- !has_unit &
    lo_v == round(lo_v) & hi_v == round(hi_v) &
    lo_v >= 1990 & hi_v <= 2100
  keep <- !is_year
  
  if (!any(keep)) return(no_number())
  
  flags <- character(0)
  
  if (any(keep & has_unit)) {
    # Explicit units win. Stray unit-less numbers (e.g. "18 MW + 6") are
    # ignored rather than guessed, and flagged.
    dropped <- keep & !has_unit
    if (any(dropped)) {
      flags <- c(flags, paste0(
        "Ignored number(s) without a unit: ", paste(lo_v[dropped], collapse = ", ")
      ))
    }
    keep <- keep & has_unit
  } else if (str_detect(s, "[A-Za-z]")) {
    flags <- c(flags, "No unit stated - assumed MW")
  }
  
  lo_mw <- lo_v * cap_unit_factor(coalesce(u1, u2))
  hi_mw <- hi_v * cap_unit_factor(coalesce(u2, u1))
  tok_min <- pmin(lo_mw, hi_mw)[keep]
  tok_max <- pmax(lo_mw, hi_mw)[keep]
  n_tok <- sum(keep)
  
  strong <- str_extract(s, CAP_AMBIG_STRONG_RE)
  if (!is.na(strong)) {
    flags <- c(flags, paste0("Ambiguous wording ('", strong, "') - check the value"))
  } else if (n_tok > 1 && !is.na(str_extract(s, CAP_AMBIG_MULTI_RE))) {
    flags <- c(flags, "Numbers may overlap (total / incl. / bracketed) - check the value")
  }
  
  mn <- sum(tok_min)
  mx <- sum(tok_max)
  
  if (mx >= 5000) flags <- c(flags, "Very large (5,000+ MW) - check the units")
  
  list(
    min = mn,
    max = mx,
    est = (mn + mx) / 2,
    flag = if (length(flags) > 0) paste(flags, collapse = "; ") else NA_character_
  )
}

# Vectorised version: one row per input value
parse_capacity_detail <- function(x) {
  res <- lapply(as.character(x), parse_capacity_one)
  
  tibble(
    MW_min = vapply(res, function(r) r$min, numeric(1)),
    MW_max = vapply(res, function(r) r$max, numeric(1)),
    MW_est = vapply(res, function(r) r$est, numeric(1)),
    Flag   = vapply(res, function(r) r$flag, character(1))
  )
}

# Kept for convenience: just the estimate
parse_capacity_mw <- function(x) parse_capacity_detail(x)$MW_est

# "$150-$180/kW" -> 165 | "$200" -> 200 | blank -> NA
parse_price <- function(x) {
  vapply(as.character(x), function(v) {
    if (is.na(v)) return(NA_real_)
    nums <- str_extract_all(str_replace_all(v, ",", ""), "[0-9]+\\.?[0-9]*")[[1]]
    if (length(nums) == 0) NA_real_ else mean(as.numeric(nums))
  }, numeric(1), USE.NAMES = FALSE)
}

# ============================================================
# SEARCH HELPERS
#   equinix texas     -> every word must match somewhere in the row
#   "digital realty"  -> exact phrase
# ============================================================

parse_search_terms <- function(term) {
  term <- str_to_lower(term)
  
  m <- str_match_all(term, '"([^"]*)"?')[[1]]
  phrases <- if (nrow(m) > 0) trimws(m[, 2]) else character(0)
  phrases <- phrases[nzchar(phrases)]
  
  rest <- trimws(str_replace_all(term, '"[^"]*"?', " "))
  words <- if (nzchar(rest)) str_split(rest, "\\s+")[[1]] else character(0)
  words <- words[nzchar(words)]
  
  c(phrases, words)
}

# Zoom a leaflet map (or proxy) to a set of points.
fit_points <- function(map, lat, lng) {
  ok <- is.finite(lat) & is.finite(lng)
  lat <- lat[ok]
  lng <- lng[ok]
  
  if (length(lat) == 0) return(map)
  
  if (min(lat) == max(lat) && min(lng) == max(lng)) {
    map %>% setView(lng = lng[1], lat = lat[1], zoom = 9)
  } else {
    map %>% fitBounds(min(lng), min(lat), max(lng), max(lat))
  }
}

# ============================================================
# DATABASE HELPERS
# ============================================================

get_con <- function() {
  dbConnect(RSQLite::SQLite(), DB_PATH)
}

initialize_database <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_meta" %in% dbListTables(con)) {
    dbWriteTable(
      con, "dc_meta",
      tibble(
        version = integer(), timestamp = character(), note = character(),
        n_rows = integer(), n_pipeline_rows = integer()
      ),
      overwrite = TRUE
    )
  }
  
  if (!"dc_current" %in% dbListTables(con)) {
    dbWriteTable(
      con, "dc_current",
      tibble(
        Operator = character(), City = character(), State = character(),
        Country = character(), Region = character(), Capacity = character(),
        Capacity_MW_est = double(), Capacity_MW_min = double(),
        Capacity_MW_max = double(), Capacity_Flag = character(),
        Utility_Rate = character(),
        Price_per_kW = character(), Cooling = character(), PUE = double(),
        Tax_Incentives = character(), Deal_Reg = character(),
        Notes = character(), Contacts = character(), is_us = logical(),
        City_clean = character(), Latitude = double(), Longitude = double()
      ),
      overwrite = TRUE
    )
  }
  
  if (!"dc_pipeline_current" %in% dbListTables(con)) {
    dbWriteTable(
      con, "dc_pipeline_current",
      tibble(
        Operator = character(), City_clean = character(), State = character(),
        Country = character(), Region = character(), is_us = logical(),
        Quarter = character(), Quarter_Date = character(),
        MW_available = double(), MW_min = double(), MW_max = double(),
        Latitude = double(), Longitude = double()
      ),
      overwrite = TRUE
    )
  }
  
  if (!"geocode_cache" %in% dbListTables(con)) {
    dbWriteTable(
      con, "geocode_cache",
      tibble(key = character(), Latitude = double(), Longitude = double()),
      overwrite = TRUE
    )
  }
}

initialize_database()

# ============================================================
# MASTER EXCEL CLEANING
# ============================================================

clean_master_df <- function(raw) {
  names(raw) <- str_squish(names(raw))
  
  # ---- Quarter columns come from the sheet itself, sorted by date ----
  q_idx <- str_detect(names(raw), regex(QUARTER_PATTERN, ignore_case = TRUE))
  names(raw)[q_idx] <- str_to_upper(names(raw)[q_idx])
  quarter_cols <- names(raw)[q_idx]
  quarter_cols <- quarter_cols[order(quarter_to_date(quarter_cols))]
  
  required_cols <- c("DC Operator", "Country", "Total Capacity (MW)")
  missing_req <- setdiff(required_cols, names(raw))
  
  if (length(missing_req) > 0) {
    stop(paste(
      "Master file is missing required columns:",
      paste(missing_req, collapse = ", ")
    ))
  }
  
  for (col in MASTER_OPTIONAL_COLS) {
    if (!col %in% names(raw)) raw[[col]] <- NA_character_
  }
  
  raw <- raw %>%
    mutate(across(everything(), ~ na_if(str_trim(as.character(.x)), ""))) %>%
    mutate(row_id = row_number())
  
  # ---- Current capacity: min / max / estimate + review flag ----
  cap <- parse_capacity_detail(raw$`Total Capacity (MW)`)
  
  # Market = what users see/filter by. City = real city, used only for geocoding.
  current <- raw %>%
    transmute(
      row_id,
      Operator = `DC Operator`,
      Market = Market,
      City = City,
      Geo_City = coalesce(City, Market),
      Address = Address,
      State = `State/Province`,
      Country = Country,
      Region = Region,
      Capacity = `Total Capacity (MW)`,
      Capacity_MW_est = cap$MW_est,
      Capacity_MW_min = cap$MW_min,
      Capacity_MW_max = cap$MW_max,
      Capacity_Flag = cap$Flag,
      Utility_Rate = `Utility Rate ($/kWh)`,
      Price_per_kW = `Price ($/kW)`,
      Cooling = Cooling,
      PUE = suppressWarnings(as.numeric(PUE)),
      Tax_Incentives = `Tax Incentives`,
      Deal_Reg = `Deal Reg`,
      Notes = Notes,
      Contacts = Contacts,
      is_us = Country == "United States",
      City_clean = coalesce(Market, City, paste0(Country, " (no city provided)"))
    )
  
  flags_current <- tibble(
    Excel_Row = raw$row_id + 1L,
    Operator = raw$`DC Operator`,
    Market = coalesce(raw$Market, raw$City, raw$Country),
    Field = "Total Capacity (MW)",
    Value_in_sheet = raw$`Total Capacity (MW)`,
    Read_as = mw_range_txt(cap$MW_est, cap$MW_min, cap$MW_max),
    Issue = cap$Flag
  ) %>%
    filter(!is.na(Issue))
  
  # ---- Upcoming capacity (one row per site x quarter) ----
  empty_pipeline <- tibble(
    row_id = integer(), Operator = character(), City_clean = character(),
    State = character(), Country = character(), Region = character(),
    is_us = logical(), Quarter = character(), Quarter_Date = as.Date(character()),
    MW_available = double(), MW_min = double(), MW_max = double()
  )
  
  pipeline <- empty_pipeline
  flags_pipeline <- tibble(
    Excel_Row = integer(), Operator = character(), Market = character(),
    Field = character(), Value_in_sheet = character(),
    Read_as = character(), Issue = character()
  )
  
  if (length(quarter_cols) > 0) {
    pipeline_meta <- raw %>%
      transmute(
        row_id,
        Operator = `DC Operator`,
        City = City,
        State = `State/Province`,
        Country = Country,
        Region = Region,
        is_us = Country == "United States",
        City_clean = coalesce(Market, City, paste0(Country, " (no city provided)"))
      )
    
    pipeline_cells <- bind_cols(pipeline_meta, raw %>% select(all_of(quarter_cols))) %>%
      pivot_longer(
        cols = all_of(quarter_cols),
        names_to = "Quarter",
        values_to = "raw_val"
      ) %>%
      filter(!is.na(raw_val))
    
    pq <- parse_capacity_detail(pipeline_cells$raw_val)
    
    pipeline_cells <- pipeline_cells %>%
      mutate(
        MW_available = pq$MW_est,
        MW_min = pq$MW_min,
        MW_max = pq$MW_max,
        Cap_Flag = pq$Flag,
        Quarter_Date = quarter_to_date(Quarter)
      )
    
    flags_pipeline <- pipeline_cells %>%
      filter(!is.na(Cap_Flag)) %>%
      transmute(
        Excel_Row = row_id + 1L,
        Operator,
        Market = City_clean,
        Field = Quarter,
        Value_in_sheet = raw_val,
        Read_as = mw_range_txt(MW_available, MW_min, MW_max),
        Issue = Cap_Flag
      )
    
    pipeline <- pipeline_cells %>%
      filter(!is.na(MW_available)) %>%
      select(
        row_id, Operator, City_clean, State, Country, Region,
        is_us, Quarter, Quarter_Date, MW_available, MW_min, MW_max
      )
  }
  
  flags <- bind_rows(flags_current, flags_pipeline) %>%
    arrange(Excel_Row)
  
  list(current = current, pipeline = pipeline, flags = flags, quarters = quarter_cols)
}

# ============================================================
# GEOCODING
# ============================================================

load_geocode_cache <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"geocode_cache" %in% dbListTables(con)) {
    return(tibble(key = character(), Latitude = double(), Longitude = double()))
  }
  
  dbReadTable(con, "geocode_cache") %>% as_tibble()
}

save_geocode_cache <- function(cache_df) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  dbWriteTable(con, "geocode_cache", cache_df, overwrite = TRUE)
}

# Real city used for geocoding (falls back to Market, then to old City_clean data)
geo_city <- function(df) {
  gc   <- if ("Geo_City" %in% names(df)) df$Geo_City else rep(NA_character_, nrow(df))
  city <- if ("City" %in% names(df)) df$City else rep(NA_character_, nrow(df))
  coalesce(gc, city)
}

# Address rows get their own cache key; all others stay city-level.
geo_key <- function(df) {
  addr <- if ("Address" %in% names(df)) df$Address else rep(NA_character_, nrow(df))
  gc <- geo_city(df)
  ifelse(
    !is.na(addr),
    paste("ADDR", addr, gc, df$State, df$Country, sep = "|"),
    paste(gc, df$State, df$Country, sep = "|")
  )
}

geocode_current <- function(current_df, existing_current = NULL, progress_fn = NULL) {
  if (!"Address" %in% names(current_df)) current_df$Address <- NA_character_
  current_df$Geo_City <- geo_city(current_df)
  cache <- load_geocode_cache()
  
  if (!is.null(existing_current) &&
      all(c("State", "Country", "Latitude", "Longitude") %in% names(existing_current)) &&
      any(c("Geo_City", "City") %in% names(existing_current))) {
    existing_coords <- existing_current %>%
      filter(!is.na(Latitude), !is.na(Longitude)) %>%
      mutate(key = geo_key(.), Latitude = as.numeric(Latitude), Longitude = as.numeric(Longitude)) %>%
      select(key, Latitude, Longitude) %>%
      distinct(key, .keep_all = TRUE)
    cache <- bind_rows(cache, existing_coords) %>% distinct(key, .keep_all = TRUE)
  }
  
  # Hand-verified coordinates always win
  cache <- bind_rows(MANUAL_COORDS, cache) %>% distinct(key, .keep_all = TRUE)
  
  to_geocode <- current_df %>%
    filter(!is.na(Geo_City)) %>%
    mutate(key = geo_key(.)) %>%
    distinct(key, .keep_all = TRUE) %>%
    filter(!key %in% cache$key)
  
  new_rows <- tibble()
  if (nrow(to_geocode) > 0 && requireNamespace("tidygeocoder", quietly = TRUE)) {
    for (i in seq_len(nrow(to_geocode))) {
      r <- to_geocode[i, ]
      addr <- paste(na.omit(c(r$Address, r$Geo_City, r$State, r$Country)), collapse = ", ")
      if (!is.null(progress_fn)) progress_fn(i, nrow(to_geocode), addr)
      
      res <- tryCatch(tidygeocoder::geo(address = addr, method = "osm", quiet = TRUE),
                      error = function(e) NULL)
      if (!is.null(res) && nrow(res) > 0 && !is.na(res$lat[1]) && !is.na(res$long[1])) {
        new_rows <- bind_rows(new_rows, tibble(
          key = r$key, Latitude = as.numeric(res$lat[1]), Longitude = as.numeric(res$long[1])))
        cat(sprintf("  -> %s : %.4f, %.4f\n", addr, res$lat[1], res$long[1]))
      }
      Sys.sleep(1)
    }
  }
  
  if (nrow(new_rows) > 0) {
    cache <- bind_rows(cache, new_rows) %>% distinct(key, .keep_all = TRUE)
    save_geocode_cache(cache)
  }
  
  current_df %>%
    mutate(key = geo_key(.), city_key = paste(Geo_City, State, Country, sep = "|")) %>%
    left_join(cache %>% select(key, Latitude, Longitude), by = "key") %>%
    left_join(cache %>% select(city_key = key, City_Lat = Latitude, City_Lon = Longitude),
              by = "city_key") %>%
    left_join(country_centroids %>% rename(Country_Lat = Latitude, Country_Lon = Longitude),
              by = "Country") %>%
    mutate(
      Latitude = case_when(
        !is.na(Latitude) ~ Latitude,
        !is.na(City_Lat) ~ City_Lat,
        is.na(Geo_City) & !is.na(Country_Lat) ~ Country_Lat,
        TRUE ~ NA_real_),
      Longitude = case_when(
        !is.na(Longitude) ~ Longitude,
        !is.na(City_Lon) ~ City_Lon,
        is.na(Geo_City) & !is.na(Country_Lon) ~ Country_Lon,
        TRUE ~ NA_real_)
    ) %>%
    select(-key, -city_key, -City_Lat, -City_Lon, -Country_Lat, -Country_Lon)
}

# ============================================================
# DATABASE READ FUNCTIONS
# ============================================================

load_current <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  df <- dbReadTable(con, "dc_current") %>% as_tibble()
  
  if ("is_us" %in% names(df)) df$is_us <- as.logical(df$is_us) else df$is_us <- FALSE
  
  defaults <- list(
    Operator = NA_character_, Market = NA_character_, Address = NA_character_,
    City = NA_character_, Geo_City = NA_character_,
    State = NA_character_, Country = NA_character_, Region = NA_character_,
    Capacity = NA_character_, Capacity_MW_est = NA_real_,
    Capacity_MW_min = NA_real_, Capacity_MW_max = NA_real_,
    Capacity_Flag = NA_character_,
    Utility_Rate = NA_character_, Price_per_kW = NA_character_,
    Cooling = NA_character_, PUE = NA_real_, Tax_Incentives = NA_character_,
    Deal_Reg = NA_character_, Notes = NA_character_, Contacts = NA_character_,
    City_clean = NA_character_, Latitude = NA_real_, Longitude = NA_real_
  )
  
  for (col in names(defaults)) {
    if (!col %in% names(df)) df[[col]] <- defaults[[col]]
  }
  
  # Older snapshots have no min/max: fall back to the single estimate
  df$Capacity_MW_min <- coalesce(df$Capacity_MW_min, df$Capacity_MW_est)
  df$Capacity_MW_max <- coalesce(df$Capacity_MW_max, df$Capacity_MW_est)
  
  df
}

empty_pipeline_tbl <- function() {
  tibble(
    Operator = character(), City_clean = character(), State = character(),
    Country = character(), Region = character(), is_us = logical(),
    Quarter = character(), Quarter_Date = as.Date(character()),
    MW_available = double(), MW_min = double(), MW_max = double(),
    Latitude = double(), Longitude = double()
  )
}

load_current_pipeline <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_pipeline_current" %in% dbListTables(con)) return(empty_pipeline_tbl())
  
  df <- dbReadTable(con, "dc_pipeline_current") %>% as_tibble()
  
  if (nrow(df) == 0) return(empty_pipeline_tbl())
  
  # SQLite may hand dates back as numbers (days since 1970) or as text
  if ("Quarter_Date" %in% names(df)) {
    qd <- df$Quarter_Date
    df$Quarter_Date <- if (is.numeric(qd)) as.Date(qd, origin = "1970-01-01") else as.Date(qd)
  } else {
    df$Quarter_Date <- quarter_to_date(df$Quarter)
  }
  
  df$is_us <- as.logical(df$is_us)
  
  if (!"Latitude" %in% names(df)) df$Latitude <- NA_real_
  if (!"Longitude" %in% names(df)) df$Longitude <- NA_real_
  
  # Older snapshots have no min/max: fall back to the single estimate
  if (!"MW_min" %in% names(df)) df$MW_min <- df$MW_available
  if (!"MW_max" %in% names(df)) df$MW_max <- df$MW_available
  df$MW_min <- coalesce(df$MW_min, df$MW_available)
  df$MW_max <- coalesce(df$MW_max, df$MW_available)
  
  df
}

list_versions <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_meta" %in% dbListTables(con)) return(tibble())
  
  dbReadTable(con, "dc_meta") %>% arrange(desc(version)) %>% as_tibble()
}

# ---------- Data-quality flags from the last import ----------

empty_flags_tbl <- function() {
  tibble(
    Excel_Row = integer(), Operator = character(), Market = character(),
    Field = character(), Value_in_sheet = character(),
    Read_as = character(), Issue = character()
  )
}

save_import_flags <- function(flags) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  dbWriteTable(con, "import_flags", flags, overwrite = TRUE)
}

load_import_flags <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"import_flags" %in% dbListTables(con)) return(empty_flags_tbl())
  
  dbReadTable(con, "import_flags") %>% as_tibble()
}

# ============================================================
# SAVE / RESTORE VERSION
# ============================================================

MAX_VERSIONS <- 15L

prune_history <- function(con, meta, keep = MAX_VERSIONS) {
  meta <- meta %>% arrange(desc(version))
  if (nrow(meta) <= keep) return(meta)
  
  drop_v <- as.integer(meta$version[(keep + 1):nrow(meta)])
  for (v in drop_v) {
    dbExecute(con, sprintf('DROP TABLE IF EXISTS "dc_history_v%d"', v))
    dbExecute(con, sprintf('DROP TABLE IF EXISTS "dc_pipeline_history_v%d"', v))
  }
  meta[seq_len(keep), ]
}

save_new_version <- function(new_current, new_pipeline, note = "Manual update") {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  meta <- if ("dc_meta" %in% dbListTables(con)) {
    dbReadTable(con, "dc_meta") %>% as_tibble()
  } else {
    tibble(version = integer(), timestamp = character(), note = character(),
           n_rows = integer(), n_pipeline_rows = integer())
  }
  
  if (!"n_pipeline_rows" %in% names(meta)) meta$n_pipeline_rows <- NA_integer_
  
  next_v <- if (nrow(meta) == 0) 1L else max(meta$version) + 1L
  
  # Snapshot whatever is currently live.
  if ("dc_current" %in% dbListTables(con)) {
    old_current <- dbReadTable(con, "dc_current")
    
    old_pipeline <- if ("dc_pipeline_current" %in% dbListTables(con)) {
      dbReadTable(con, "dc_pipeline_current")
    } else {
      tibble()
    }
    
    dbWriteTable(con, paste0("dc_history_v", next_v), old_current, overwrite = TRUE)
    dbWriteTable(con, paste0("dc_pipeline_history_v", next_v), old_pipeline, overwrite = TRUE)
    
    meta <- bind_rows(
      meta,
      tibble(
        version = next_v,
        timestamp = as.character(Sys.time()),
        note = paste0("Auto-snapshot before: ", note),
        n_rows = nrow(old_current),
        n_pipeline_rows = nrow(old_pipeline)
      )
    )
    
    next_v <- next_v + 1L
  }
  
  dbWriteTable(con, "dc_current", new_current, overwrite = TRUE)
  dbWriteTable(con, "dc_pipeline_current", new_pipeline, overwrite = TRUE)
  
  dbWriteTable(con, paste0("dc_history_v", next_v), new_current, overwrite = TRUE)
  dbWriteTable(con, paste0("dc_pipeline_history_v", next_v), new_pipeline, overwrite = TRUE)
  
  meta <- bind_rows(
    meta,
    tibble(
      version = next_v,
      timestamp = as.character(Sys.time()),
      note = note,
      n_rows = nrow(new_current),
      n_pipeline_rows = nrow(new_pipeline)
    )
  )
  
  meta <- prune_history(con, meta)
  dbWriteTable(con, "dc_meta", meta, overwrite = TRUE)
  dbExecute(con, "VACUUM")
}

restore_version <- function(v) {
  con <- get_con()
  
  tbl_name <- paste0("dc_history_v", v)
  
  if (!tbl_name %in% dbListTables(con)) {
    dbDisconnect(con)
    return(FALSE)
  }
  
  old_current <- dbReadTable(con, tbl_name)
  pipe_tbl <- paste0("dc_pipeline_history_v", v)
  
  old_pipeline <- if (pipe_tbl %in% dbListTables(con)) dbReadTable(con, pipe_tbl) else tibble()
  
  dbDisconnect(con)
  
  save_new_version(old_current, old_pipeline, note = paste("Restored from version", v))
  
  TRUE
}

# ============================================================
# IMPORT CHANGE DETECTION
# ============================================================

# File hash + schema version: changing the parsing logic (bump
# IMPORT_SCHEMA_VERSION) forces one fresh import on the next launch.
file_signature <- function() {
  paste0(unname(tools::md5sum(MASTER_FILE)), "|schema", IMPORT_SCHEMA_VERSION)
}

get_stored_signature <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"import_state" %in% dbListTables(con)) return(NA_character_)
  
  df <- dbReadTable(con, "import_state")
  if (nrow(df) == 0) NA_character_ else df$signature[1]
}

set_stored_signature <- function(sig) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  dbWriteTable(
    con, "import_state",
    tibble(signature = sig, imported_at = as.character(Sys.time())),
    overwrite = TRUE
  )
}

# ============================================================
# IMPORT DC.XLSX
# ============================================================

import_master_excel <- function(progress_fn = NULL) {
  if (!file.exists(MASTER_FILE)) {
    stop(paste0(
      "Could not find ", MASTER_FILE,
      ". Make sure it is in the same folder as app.R."
    ))
  }
  
  # Hash before reading so a mid-import edit isn't mistaken for "already imported".
  import_sig <- file_signature()
  
  raw <- readxl::read_excel(MASTER_FILE, col_types = "text")
  
  parsed <- clean_master_df(raw)
  
  existing_current <- tryCatch(load_current(), error = function(e) NULL)
  
  parsed$current <- geocode_current(
    parsed$current,
    existing_current = existing_current,
    progress_fn = progress_fn
  )
  
  # Each upcoming-capacity entry inherits the coordinates of its own sheet row
  parsed$pipeline <- parsed$pipeline %>%
    left_join(parsed$current %>% select(row_id, Latitude, Longitude), by = "row_id") %>%
    select(-row_id)
  
  parsed$current <- parsed$current %>% select(-row_id)
  
  save_new_version(parsed$current, parsed$pipeline, note = "Refresh from DC.xlsx")
  save_import_flags(parsed$flags)
  set_stored_signature(import_sig)
  
  list(
    current = parsed$current,
    pipeline = parsed$pipeline,
    flags = parsed$flags,
    quarters = parsed$quarters
  )
}

# ============================================================
# STARTUP IMPORT (only if DC.xlsx changed since the last import)
# ============================================================

if (!file.exists(MASTER_FILE)) {
  cat(MASTER_FILE, "not found at startup - using existing SQLite data only.\n")
  
} else if (identical(file_signature(), get_stored_signature())) {
  cat(MASTER_FILE, "unchanged since last import - skipping startup import.\n")
  
} else {
  cat("Importing", MASTER_FILE, "at startup...\n")
  flush.console()
  
  startup_import_result <- tryCatch({
    import_master_excel(
      progress_fn = function(i, n, addr) {
        cat(sprintf(" Geocoding %d/%d: %s\n", i, n, addr))
        flush.console()
      }
    )
  }, error = function(e) {
    cat("Startup import FAILED:", conditionMessage(e), "\n")
    cat("App will start with existing SQLite data instead.\n")
    NULL
  })
  
  if (!is.null(startup_import_result)) {
    cat(sprintf(
      "Startup import complete: %d sites, %d upcoming-capacity entries, %d quarter columns.\n",
      nrow(startup_import_result$current),
      nrow(startup_import_result$pipeline),
      length(startup_import_result$quarters)
    ))
    
    fl <- startup_import_result$flags
    if (nrow(fl) > 0) {
      cat(sprintf("%d capacity entries need review (see Version History > Data quality check):\n", nrow(fl)))
      for (i in seq_len(min(nrow(fl), 10))) {
        cat(sprintf("  row %d | %s | %s | '%s' | %s\n",
                    fl$Excel_Row[i], fl$Operator[i], fl$Field[i],
                    fl$Value_in_sheet[i], fl$Issue[i]))
      }
      if (nrow(fl) > 10) cat(sprintf("  ... and %d more\n", nrow(fl) - 10))
    }
  }
}

# ============================================================
# CLIENT PDF (Upcoming Capacity, operator names hidden)
#   Needs: install.packages("pagedown") + Google Chrome installed.
#   Logo:  www/logo.png, shown on the dark header and footer so a
#          light/white logo stays visible.
# ============================================================

CLIENT_PDF_EYEBROW  <- "Capacity Brief"
CLIENT_PDF_TITLE    <- "Upcoming Data Center Capacity"
CLIENT_PDF_SUBTITLE <- "Near-term capacity coming available across the ByteBridge partner network"
CLIENT_PDF_CTA_HEAD <- "Interested in a site?"
CLIENT_PDF_CONTACT  <- "sales@bytebt.com"
CLIENT_PDF_DISCLAIMER <- paste(
  "Confidential — for discussion purposes only.",
  "Capacity and delivery dates are indicative, subject to change",
  "and must be confirmed with ByteBridge."
)
CLIENT_PDF_LOGO <- "www/logo.png"

client_pdf_logo <- function() {
  f <- CLIENT_PDF_LOGO
  if (!file.exists(f)) return(NULL)
  mime <- if (grepl("\\.svg$", f, ignore.case = TRUE)) "image/svg+xml"
          else if (grepl("\\.jpe?g$", f, ignore.case = TRUE)) "image/jpeg"
          else "image/png"
  paste0("data:", mime, ";base64,",
         jsonlite::base64_enc(readBin(f, "raw", file.info(f)$size)))
}

CLIENT_PDF_CSS <- "
@page { size: Letter landscape; margin: 0.4in 0.45in 0.7in 0.45in; }
* { box-sizing: border-box; }
html, body {
  margin: 0; padding: 0; color: #142233;
  font-family: 'Inter', 'Helvetica Neue', Helvetica, Arial, sans-serif;
  font-feature-settings: 'tnum' 1;
  -webkit-print-color-adjust: exact; print-color-adjust: exact;
}

/* ---------- Hero ---------- */
.hero {
  position: relative; overflow: hidden; border-radius: 16px;
  padding: 22px 28px 24px; color: #fff;
  background:
    radial-gradient(520px 260px at 92% -20%, rgba(25,181,201,.35), transparent 60%),
    radial-gradient(420px 240px at 70% 130%, rgba(47,124,246,.30), transparent 60%),
    linear-gradient(120deg, #07182A 0%, #0D2742 55%, #123A63 100%);
}
.hero::after {
  content: ''; position: absolute; left: 0; right: 0; bottom: 0; height: 4px;
  background: linear-gradient(90deg, #2F7CF6, #19B5C9);
}
.hero-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px; }
.hero-logo { height: 38px; display: block; }
.wordmark { font-size: 20px; font-weight: 800; letter-spacing: -.01em; }
.hero-date { text-align: right; font-size: 9px; letter-spacing: .14em; text-transform: uppercase; color: #8FB3D9; }
.hero-date b { display: block; margin-top: 3px; font-size: 13px; letter-spacing: 0; text-transform: none; color: #fff; font-weight: 600; }
.eyebrow {
  display: inline-block; font-size: 9px; font-weight: 700; letter-spacing: .18em; text-transform: uppercase;
  color: #7FE3EF; border: 1px solid rgba(127,227,239,.45); border-radius: 999px; padding: 3px 10px; margin-bottom: 10px;
}
.hero h1 { margin: 0; font-size: 30px; font-weight: 800; letter-spacing: -.02em; line-height: 1.1; }
.hero .sub { margin-top: 6px; font-size: 12.5px; color: #B9CCE0; }

/* ---------- KPI cards ---------- */
.kpis { display: flex; gap: 12px; margin: 16px 0 20px; }
.kpi {
  flex: 1; position: relative; overflow: hidden;
  border: 1px solid #E3E9F0; border-radius: 12px; padding: 12px 16px 12px 18px; background: #fff;
}
.kpi::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
               background: linear-gradient(180deg, #2F7CF6, #19B5C9); }
.kpi .k { font-size: 8.5px; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; color: #64748B; }
.kpi .v { margin-top: 4px; font-size: 24px; font-weight: 800; letter-spacing: -.02em; color: #0D2742; }
.kpi .v small { font-size: 12px; font-weight: 600; color: #64748B; margin-left: 3px; }

/* ---------- Section title ---------- */
.sec { display: flex; align-items: baseline; justify-content: space-between; margin: 0 2px 10px; }
.sec h2 { margin: 0; font-size: 14px; font-weight: 750; color: #0D2742; }
.sec span { font-size: 9.5px; color: #64748B; }

/* ---------- Quarter sections (the big break) ---------- */
.qsec { margin-top: 26px; }
.qsec.first { margin-top: 0; }
.qhead {
  display: flex; align-items: center; justify-content: space-between;
  padding: 13px 20px 12px; border-radius: 12px 12px 0 0; color: #fff;
  background:
    radial-gradient(380px 120px at 95% 0%, rgba(25,181,201,.30), transparent 70%),
    linear-gradient(110deg, #07182A 0%, #0D2742 60%, #123A63 100%);
  border-bottom: 3px solid #19B5C9;
  break-inside: avoid; break-after: avoid; page-break-after: avoid;
}
.qk { font-size: 8px; font-weight: 800; letter-spacing: .2em; text-transform: uppercase; color: #7FE3EF; }
.qt { margin-top: 2px; font-size: 24px; font-weight: 800; letter-spacing: -.02em; line-height: 1.05; }
.qstats { display: flex; gap: 26px; text-align: right; }
.qstats b { display: block; font-size: 17px; font-weight: 800; letter-spacing: -.01em; }
.qstats span { font-size: 7.5px; font-weight: 700; letter-spacing: .16em; text-transform: uppercase; color: #8FB3D9; }

/* ---------- Table ---------- */
table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 10.5px;
        border: 1px solid #E3E9F0; border-top: none; border-radius: 0 0 12px 12px; }
thead { display: table-header-group; }
th {
  background: #F3F6FA; color: #64748B; font-size: 8px; font-weight: 700; letter-spacing: .14em;
  text-transform: uppercase; text-align: left; padding: 8px 14px; border-bottom: 1px solid #E3E9F0;
}
th:first-child { text-align: center; }
th .thq { color: #1D4ED8; }
td { padding: 8px 14px; border-bottom: 1px solid #EEF2F6; vertical-align: middle; }
tbody tr:last-child td { border-bottom: none; }
td.n { text-align: center; color: #94A3B8; font-weight: 600; }
tr { break-inside: avoid; page-break-inside: avoid; }
.loc { padding-left: 26px; }
.loc b { display: block; font-size: 11.5px; font-weight: 650; color: #0D2742; }
.loc span { display: block; margin-top: 1px; font-size: 9.5px; color: #64748B; }
.cap { display: flex; align-items: center; gap: 12px; }
.bar { flex: 1; height: 7px; border-radius: 999px; background: #EDF2F7; overflow: hidden; }
.bar > i { display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #2F7CF6, #19B5C9); }
.cap b { min-width: 44px; text-align: right; font-size: 12px; font-weight: 750; color: #0D2742; }

/* ---------- Region sub-groups inside each quarter ---------- */
tr.rg td { padding: 9px 14px 8px; border-bottom: 1px solid #E3E9F0; break-after: avoid; page-break-after: avoid; }
.rh { display: flex; align-items: center; justify-content: space-between; }
.rl { display: flex; align-items: center; gap: 9px; }
.rdot { width: 10px; height: 10px; border-radius: 3px; display: inline-block; }
.rt { font-size: 11.5px; font-weight: 800; letter-spacing: .02em; color: #0D2742; text-transform: uppercase; }
.rm { font-size: 9.5px; font-weight: 650; color: #475569; }

/* ---------- Call to action ---------- */
.cta {
  margin-top: 24px; padding: 16px 20px; border-radius: 12px; break-inside: avoid;
  background:
    radial-gradient(360px 140px at 100% 0%, rgba(25,181,201,.28), transparent 70%),
    linear-gradient(120deg, #0D2742, #123A63);
  color: #fff; display: flex; align-items: center; justify-content: space-between; gap: 20px;
}
.cta-text b { display: block; font-size: 22px; font-weight: 800; letter-spacing: -.01em; }
.cta-btn {
  flex: none; display: inline-block; padding: 12px 26px; border-radius: 999px; text-decoration: none;
  background: #fff; color: #0D2742; font-size: 22px; font-weight: 800; letter-spacing: -.01em;
}
.cta-btn small { display: block; font-size: 8.5px; font-weight: 700; letter-spacing: .18em;
                 text-transform: uppercase; color: #19A3B5; margin-bottom: 2px; }
"

# Removes any operator name from a text vector (e.g. a Market called
# "Equinix Dallas" becomes "Dallas")
scrub_operators <- function(x, operators) {
  ops <- unique(operators[!is.na(operators) & nchar(operators) >= 2])
  ops <- ops[order(-nchar(ops))]            # longest names first
  esc <- function(s) gsub("([][{}()+*^$|\\\\?.])", "\\\\\\1", s)
  for (op in ops) {
    x <- gsub(paste0("\\b", esc(op), "\\b"), "", x, ignore.case = TRUE, perl = TRUE)
  }
  x <- str_squish(gsub("^[[:space:][:punct:]]+|[[:space:][:punct:]]+$", "", x))
  ifelse(is.na(x), "", x)
}

# ---------- Client regions ----------
# USA sites use the Region you set in DC.xlsx ("South West" -> "USA – South West").
# Only if a US row has no Region does the state list below fill it in.
# Everything else rolls up to Canada / Latin America / EMEA / APAC by country.

CLIENT_PDF_US_REGIONS <- list(           # fallback only, for US rows with a blank Region
  "North West" = c("WA", "OR", "ID", "MT", "WY", "AK"),
  "South West" = c("CA", "NV", "AZ", "UT", "CO", "NM", "HI", "TX", "OK"),
  "Midwest"    = c("ND", "SD", "NE", "KS", "MN", "IA", "MO", "WI", "IL", "IN", "MI", "OH"),
  "North East" = c("ME", "NH", "VT", "MA", "RI", "CT", "NY", "NJ", "PA", "DE", "MD", "DC"),
  "South East" = c("VA", "WV", "NC", "SC", "GA", "FL", "AL", "MS", "TN", "KY", "AR", "LA")
)

CLIENT_PDF_WORLD_REGIONS <- list(
  "Canada" = c("Canada"),
  "Latin America" = c("Mexico", "Brazil", "Chile", "Colombia", "Argentina", "Uruguay", "Peru",
               "Panama", "Costa Rica", "Ecuador", "Dominican Republic", "Guatemala"),
  "EMEA"   = c("United Kingdom", "UK", "Ireland", "Germany", "France", "Netherlands", "Spain",
               "Italy", "Sweden", "Norway", "Sweden/Norway", "Poland", "Switzerland", "Portugal",
               "Belgium", "Denmark", "Finland", "Austria", "Czech Republic", "Czechia", "Greece",
               "Iceland", "Luxembourg", "Romania", "Hungary", "Bulgaria", "Croatia", "Serbia",
               "Turkey", "Israel", "United Arab Emirates", "UAE", "Saudi Arabia", "Qatar",
               "Bahrain", "Oman", "Kuwait", "Egypt", "Morocco", "South Africa", "Kenya", "Nigeria"),
  "APAC"   = c("Australia", "New Zealand", "Japan", "South Korea", "Korea", "China", "Hong Kong",
               "Taiwan", "Singapore", "Malaysia", "Indonesia", "Thailand", "Vietnam",
               "Philippines", "India", "Bangladesh", "Sri Lanka", "Pakistan")
)

# Preferred display order. USA regions from the sheet that are not listed
# here still show, alphabetically, after the listed ones.
CLIENT_PDF_REGION_ORDER <- c(
  paste("USA –", c("North West", "South West", "Midwest", "North East", "South East")),
  "USA – Other",
  names(CLIENT_PDF_WORLD_REGIONS), "Other"
)

client_region_order <- function(regions) {
  regions <- unique(regions[!is.na(regions)])
  known <- intersect(CLIENT_PDF_REGION_ORDER, regions)
  extra <- setdiff(regions, known)
  extra_us <- sort(extra[startsWith(extra, "USA")])
  extra_other <- sort(setdiff(extra, extra_us))
  us_known <- known[startsWith(known, "USA")]
  c(setdiff(us_known, "USA – Other"), extra_us, intersect("USA – Other", us_known),
    setdiff(known, us_known), extra_other)
}

us_state_abb <- function(s) {
  s <- str_squish(coalesce(as.character(s), ""))
  up <- toupper(s)
  ab <- ifelse(up %in% c(state.abb, "DC"), up,
               state.abb[match(tolower(s), tolower(state.name))])
  ab[tolower(s) %in% c("district of columbia", "washington dc", "washington, d.c.")] <- "DC"
  ab
}

# "south west", "USA - South West", "US: Southwest" -> "South West"-style label
clean_sheet_region <- function(r) {
  r <- str_squish(coalesce(as.character(r), ""))
  r <- str_replace(r, regex("^(usa|u\\.s\\.a?\\.?|us|united states)\\s*[-–—:|/]?\\s*",
                            ignore_case = TRUE), "")
  r <- str_squish(r)
  ifelse(nzchar(r), str_to_title(r), NA_character_)
}

client_region <- function(state, country, sheet_region = NULL) {
  country <- str_squish(coalesce(as.character(country), ""))
  is_us <- tolower(country) %in% c("united states", "usa", "us", "united states of america")
  if (is.null(sheet_region)) sheet_region <- rep(NA_character_, length(country))

  # USA -> "USA – <Region from the sheet>", state list only as a fallback
  ab <- us_state_abb(state)
  us_lookup <- setNames(rep(names(CLIENT_PDF_US_REGIONS), lengths(CLIENT_PDF_US_REGIONS)),
                        unlist(CLIENT_PDF_US_REGIONS))
  us_sub <- coalesce(clean_sheet_region(sheet_region), unname(us_lookup[ab]), "Other")
  us_lbl <- paste("USA –", us_sub)

  # Rest of world
  world_lookup <- setNames(rep(names(CLIENT_PDF_WORLD_REGIONS), lengths(CLIENT_PDF_WORLD_REGIONS)),
                           tolower(unlist(CLIENT_PDF_WORLD_REGIONS)))
  world_lbl <- unname(world_lookup[tolower(country)])

  ifelse(is_us, us_lbl, coalesce(world_lbl, "Other"))
}

# Colour swatch per region family: c(dot, row tint)
client_region_colors <- function(region) {
  case_when(
    startsWith(region, "USA")  ~ "#2F7CF6|#EEF4FF",
    region == "Canada"         ~ "#7C3AED|#F4F0FE",
    region == "Latin America"  ~ "#F59E0B|#FFF7E8",
    region == "EMEA"           ~ "#10B981|#ECFAF4",
    region == "APAC"           ~ "#E11D48|#FDEEF2",
    TRUE                       ~ "#94A3B8|#F3F6FA"
  )
}

# df must already have the Operator column removed; `operators` is the
# full list of names to scrub from every remaining text field.
client_pdf_html <- function(df, operators = character(0)) {
  stopifnot(!"Operator" %in% names(df))
  e <- function(x) htmltools::htmlEscape(coalesce(as.character(x), ""))
  logo <- client_pdf_logo()
  
  df$Region_Client <- client_region(df$State, df$Country, df$Region)   # before scrubbing
  df$Region_Client <- scrub_operators(df$Region_Client, operators)
  for (col in c("City", "State", "Country", "Quarter")) {
    df[[col]] <- scrub_operators(as.character(df[[col]]), operators)
  }
  df$City <- ifelse(nzchar(df$City), df$City, "Undisclosed")
  df$MW <- coalesce(as.numeric(df$`MW Available`), 0)
  
  mw_max <- max(df$MW, 1, na.rm = TRUE)
  quarters <- unique(df$Quarter)             # already in date order
  regions_all <- client_region_order(df$Region_Client)
  
  plural <- function(n) if (n == 1) " opportunity" else " opportunities"
  n <- 0L
  
  row_html <- function(r) {
    n <<- n + 1L
    sub <- c(r$State, r$Country)
    sub <- sub[!is.na(sub) & nzchar(sub)]
    paste0(
      "<tr class='row'>",
      "<td class='n'>", n, "</td>",
      "<td class='loc'><b>", e(r$City), "</b>",
      if (length(sub)) paste0("<span>", e(paste(sub, collapse = " · ")), "</span>") else "",
      "</td>",
      "<td><div class='cap'><div class='bar'><i style='width:",
      max(4, round(100 * r$MW / mw_max)), "%'></i></div><b>", e(fmt_mw(r$MW)), "</b></div></td>",
      "</tr>"
    )
  }
  
  # Quarter -> Region -> sites
  quarter_html <- vapply(seq_along(quarters), function(qi) {
    q <- quarters[qi]
    gq <- df[df$Quarter == q, ]
    regs <- client_region_order(gq$Region_Client)
    
    body <- vapply(regs, function(rg) {
      g <- gq[gq$Region_Client == rg, ]
      g <- g[order(-g$MW), ]
      col <- strsplit(client_region_colors(rg), "|", fixed = TRUE)[[1]]
      paste0(
        "<tr class='rg'><td colspan='3' style='background:", col[2], "'><div class='rh'>",
        "<div class='rl'><span class='rdot' style='background:", col[1], "'></span>",
        "<span class='rt'>", e(rg), "</span></div>",
        "<span class='rm'>", nrow(g), plural(nrow(g)), " &middot; ", e(fmt_mw(sum(g$MW))), " MW</span>",
        "</div></td></tr>",
        paste0(vapply(seq_len(nrow(g)), function(i) row_html(g[i, ]), character(1)), collapse = "")
      )
    }, character(1))
    
    paste0(
      "<section class='qsec", if (qi == 1) " first" else "", "'>",
      "<div class='qhead'><div><div class='qk'>Coming online</div><div class='qt'>", e(q), "</div></div>",
      "<div class='qstats'>",
      "<div><b>", nrow(gq), "</b><span>", trimws(plural(nrow(gq))), "</span></div>",
      "<div><b>", length(regs), "</b><span>", if (length(regs) == 1) "region" else "regions", "</span></div>",
      "<div><b>", e(fmt_mw(sum(gq$MW))), "</b><span>MW</span></div>",
      "</div></div>",
      "<table><colgroup><col style='width:6%'><col style='width:52%'><col style='width:42%'></colgroup>",
      "<thead><tr><th>#</th><th>Market / Location <span class='thq'>&middot; ", e(q), "</span></th>",
      "<th>Capacity (MW)</th></tr></thead>",
      "<tbody>", paste0(body, collapse = ""), "</tbody></table>",
      "</section>"
    )
  }, character(1))
  
  d <- Sys.Date()
  report_date <- paste0(format(d, "%B "), as.integer(format(d, "%d")), format(d, ", %Y"))
  
  kpi <- function(k, v, unit = "") {
    paste0("<div class='kpi'><div class='k'>", k, "</div><div class='v'>", v,
           if (nzchar(unit)) paste0("<small>", unit, "</small>") else "", "</div></div>")
  }
  
  paste0(
    "<!doctype html><html><head><meta charset='utf-8'>",
    "<link rel='preconnect' href='https://fonts.googleapis.com'>",
    "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap' rel='stylesheet'>",
    "<style>", CLIENT_PDF_CSS, "</style></head><body>",
    
    "<div class='hero'>",
    "<div class='hero-top'>",
    if (!is.null(logo)) paste0("<img class='hero-logo' src='", logo, "' alt='ByteBridge'>")
    else "<div class='wordmark'>ByteBridge</div>",
    "<div class='hero-date'>Report date<b>", e(report_date), "</b></div>",
    "</div>",
    "<div class='eyebrow'>", e(CLIENT_PDF_EYEBROW), "</div>",
    "<h1>", e(CLIENT_PDF_TITLE), "</h1>",
    "<div class='sub'>", e(CLIENT_PDF_SUBTITLE), "</div>",
    "</div>",
    
    "<div class='kpis'>",
    kpi("Opportunities", nrow(df)),
    kpi("Upcoming capacity", e(fmt_mw(sum(df$MW))), "MW"),
    kpi("Earliest availability", e(quarters[1])),
    kpi("Regions", length(regions_all)),
    "</div>",
    
    "<div class='sec'><h2>Capacity pipeline</h2>",
    "<span>By quarter coming online, then by region &middot; all figures in MW</span></div>",
    paste0(quarter_html, collapse = ""),
    
    "<div class='cta'><div class='cta-text'><b>", e(CLIENT_PDF_CTA_HEAD), "</b></div>",
    "<a class='cta-btn' href='mailto:", e(CLIENT_PDF_CONTACT), "'><small>Contact us</small>",
    e(CLIENT_PDF_CONTACT), "</a></div>",
    "</body></html>"
  )
}

# Footer on every page. The logo sits on a dark pill so a white logo stays visible.
client_pdf_footer <- function() {
  logo <- client_pdf_logo()
  brand <- if (!is.null(logo)) {
    paste0("<img src='", logo, "' style='height:11px; display:block;'>")
  } else {
    "<span style='color:#fff; font-weight:700; font-size:8px;'>ByteBridge</span>"
  }

  paste0(
    "<div style='width:100%; padding:0 0.45in; display:flex; align-items:center; gap:12px;",
    " font-family:Inter,Helvetica,Arial,sans-serif; font-size:7.5px; color:#64748B;",
    " -webkit-print-color-adjust:exact; print-color-adjust:exact;'>",
    "<span style='background:#0D2742; border-radius:6px; padding:4px 8px; display:inline-flex;",
    " align-items:center; -webkit-print-color-adjust:exact;'>", brand, "</span>",
    "<span style='flex:1;'>", htmltools::htmlEscape(CLIENT_PDF_DISCLAIMER), "</span>",
    "<span style='font-weight:600; color:#0D2742;'>Page <span class='pageNumber'></span>",
    " of <span class='totalPages'></span></span>",
    "</div>"
  )
}

render_client_pdf <- function(df, out_file, operators = character(0)) {
  if (!requireNamespace("pagedown", quietly = TRUE)) {
    stop("The pagedown package is not installed. Run install.packages('pagedown').")
  }
  if (!nzchar(tryCatch(pagedown::find_chrome(), error = function(e) ""))) {
    stop("Google Chrome was not found; it is needed to build the PDF.")
  }

  tmp_html <- tempfile(fileext = ".html")
  tmp_pdf <- tempfile(fileext = ".pdf")
  con <- file(tmp_html, open = "w", encoding = "UTF-8")
  writeLines(client_pdf_html(df, operators), con)
  close(con)

  pagedown::chrome_print(
    input = tmp_html, output = tmp_pdf, wait = 2, timeout = 60, verbose = 0,
    options = list(
      displayHeaderFooter = TRUE,
      headerTemplate = "<div></div>",
      footerTemplate = client_pdf_footer(),
      preferCSSPageSize = TRUE,
      printBackground = TRUE
    )
  )
  file.copy(tmp_pdf, out_file, overwrite = TRUE)
}

# ============================================================
# UI HELPERS
# ============================================================

# Range slider with two editable number boxes (kept in sync by the server)
range_slider <- function(id, min = 0, max = 100, value = c(0, 100), step = 1) {
  div(
    class = "range-wrap",
    div(
      class = "range-inputs",
      numericInput(paste0(id, "_lo"), NULL, value = value[1], min = min, step = step),
      span(class = "range-dash", "\u2013"),
      numericInput(paste0(id, "_hi"), NULL, value = value[2], min = min, step = step)
    ),
    sliderInput(id, NULL, min = min, max = max, value = value, step = step)
  )
}

# Headline number tile
kpi <- function(tone, ic, label, out_id) {
  div(
    class = paste("kpi", paste0("kpi-", tone)),
    div(class = "kpi-icon", icon(ic)),
    div(
      class = "kpi-text",
      div(class = "kpi-label", label),
      div(class = "kpi-value", textOutput(out_id, inline = TRUE))
    )
  )
}

# Card header with an icon and an optional hint
card_title <- function(ic, title, hint = NULL) {
  card_header(
    class = "dc-card-head",
    span(class = "dc-card-ic", icon(ic)),
    span(class = "dc-card-title", title),
    if (!is.null(hint)) span(class = "dc-card-hint", hint)
  )
}

# Shared look for plotly charts on the dark theme
dark_plot <- function(p, ...) {
  p %>%
    layout(
      paper_bgcolor = "rgba(0,0,0,0)",
      plot_bgcolor = "rgba(0,0,0,0)",
      font = list(color = "#E2E8F0", family = "Inter"),
      hoverlabel = list(bgcolor = "#0B1220", bordercolor = "#334155",
                        font = list(color = "#F1F5F9", family = "Inter")),
      ...
    ) %>%
    config(displayModeBar = FALSE, responsive = TRUE)
}

empty_plot <- function(msg) {
  plot_ly() %>%
    dark_plot(
      xaxis = list(visible = FALSE),
      yaxis = list(visible = FALSE),
      annotations = list(list(
        text = msg, x = 0.5, y = 0.5, xref = "paper", yref = "paper",
        showarrow = FALSE, font = list(size = 14, color = "#94A3B8")
      ))
    )
}

# Adds an in-cell bar to a numeric column (skipped if the column has no data)
add_color_bar <- function(dt, df, col, color = "#2563EB") {
  v <- suppressWarnings(as.numeric(df[[col]]))
  rng <- suppressWarnings(range(v, na.rm = TRUE))
  if (length(v) == 0 || !all(is.finite(rng))) return(dt)
  if (rng[1] == rng[2]) rng[2] <- rng[1] + 1
  
  dt %>% formatStyle(
    col,
    background = styleColorBar(rng, color),
    backgroundSize = "96% 62%",
    backgroundRepeat = "no-repeat",
    backgroundPosition = "center"
  )
}

# ============================================================
# JAVASCRIPT
# ============================================================

APP_JS <- r"---(
$(function() {
  // Sliders created inside hidden panels can render with zero width;
  // refresh them whenever the tab changes.
  $(document).on('shown.bs.tab', function() {
    setTimeout(function() {
      $('input.js-range-slider').each(function() {
        var inst = $(this).data('ionRangeSlider');
        if (inst) inst.update();
      });
      window.dispatchEvent(new Event('resize'));
    }, 200);
  });

    // ---- Highlight window timeline (quarter strip) ----
    var hlOn = false;
    var tlDrag = false;
    function tlBars() { return $('#hl_tl_wrap .tl-col'); }
    function tlSet(idx) {
      var $b = tlBars();
      if (!$b.length) return;
      idx = Math.max(0, Math.min(idx, $b.length - 1));
      var mw = 0;
      $b.each(function(i) {
        var inside = i <= idx;
        $(this).toggleClass('in', inside).toggleClass('edge', i === idx)
               .attr('aria-pressed', i === idx ? 'true' : 'false');
        if (inside) mw += parseFloat($(this).attr('data-mw')) || 0;
      });
      var q = $b.eq(idx).attr('data-q');
      $('#hl_tl_read').html('Through <b>' + q + '</b> &middot; ' +
        Math.round(mw).toLocaleString() + ' MW arriving');
      $('#hl_tl_wrap').attr('data-idx', idx);
      Shiny.setInputValue('hl_through', q);
    }
    function tlInit() {
      var $w = $('#hl_tl_wrap');
      if (!$w.length) return;
      $w.toggleClass('is-on', hlOn);
      tlSet(parseInt($w.attr('data-default'), 10) || 0);
    }
    $(document).on('shiny:value', function(e) {
      if (e.name === 'hl_timeline') setTimeout(tlInit, 60);
    });
    $(document).on('pointerdown', '#hl_tl_wrap .tl-col', function() {
      tlDrag = true;
      tlSet($(this).index());
    });
    $(document).on('pointerenter', '#hl_tl_wrap .tl-col', function() {
      if (tlDrag) tlSet($(this).index());
    });
    $(document).on('pointerup pointercancel', function() { tlDrag = false; });
    $(document).on('click', '#hl_tl_wrap .tl-col', function() {
      tlSet($(this).index());
    });
    $(document).on('keydown', '#hl_tl_wrap .tl-col', function(e) {
      var i = $(this).index();
      if (e.key === 'ArrowRight') { e.preventDefault(); tlSet(i + 1); tlBars().eq(Math.min(i + 1, tlBars().length - 1)).focus(); }
      if (e.key === 'ArrowLeft')  { e.preventDefault(); tlSet(i - 1); tlBars().eq(Math.max(i - 1, 0)).focus(); }
    });
    $(document).on('click', '#hl_tl_wrap .tl-chip', function() {
      var n = parseInt($(this).attr('data-n'), 10);
      tlSet(n >= 99 ? tlBars().length - 1 : n - 1);
    });

    // Highlight button: off <-> on
   Shiny.addCustomMessageHandler('hl_state', function(on) {
    hlOn = on;
    $('#hl_tl_wrap').toggleClass('is-on', on);
    $('#highlight_toggle')
      .toggleClass('is-on', on)
      .attr('aria-pressed', on ? 'true' : 'false')
      .find('.hl-sub')
      .text(on ? 'Showing on map' : 'Click to highlight on map');
  });

  // ==========================================================
  // SEARCH BAR: suggestions, arrow keys, "/" or Ctrl/Cmd+K, clear
  // ==========================================================
  var SUG = [];
  var shown = [];
  var active = -1;

  function $inp() { return $('#global_search'); }
  function $box() { return $('#sb_suggest'); }

  function syncHasText() {
    var v = $inp().val() || '';
    $('.sb-search').toggleClass('has-text', v.length > 0);
  }

  function hideSuggest() {
    $box().removeClass('open').empty();
    $inp().attr('aria-expanded', 'false');
    shown = [];
    active = -1;
  }

  function highlightInto($el, text, q) {
    var i = q ? text.toLowerCase().indexOf(q) : -1;
    if (i < 0) { $el.text(text); return; }
    $el.append(document.createTextNode(text.slice(0, i)));
    $el.append($('<mark>').text(text.slice(i, i + q.length)));
    $el.append(document.createTextNode(text.slice(i + q.length)));
  }

  function renderSuggest() {
    var raw = $inp().val() || '';
    var q = raw.replace(/"/g, '').trim().toLowerCase();
    if (!q || document.activeElement !== $inp()[0]) { hideSuggest(); return; }

    var starts = [], inside = [];
    for (var i = 0; i < SUG.length; i++) {
      var p = SUG[i].t.toLowerCase().indexOf(q);
      if (p === 0) starts.push(SUG[i]);
      else if (p > 0) inside.push(SUG[i]);
      if (starts.length >= 8) break;
    }
    shown = starts.concat(inside).slice(0, 8);

    if (!shown.length) { hideSuggest(); return; }

    var $b = $box().empty();
    shown.forEach(function(s, idx) {
      var $it = $('<div class="sb-sug-item" role="option"></div>').attr('data-i', idx);
      var $t = $('<span class="sb-sug-text"></span>');
      highlightInto($t, s.t, q);
      $it.append($t, $('<span class="sb-sug-type"></span>').text(s.k));
      $b.append($it);
    });
    active = -1;
    $b.addClass('open');
    $inp().attr('aria-expanded', 'true');
  }

  function setActive(n) {
    var $items = $box().children('.sb-sug-item');
    if (!$items.length) return;
    active = (n + $items.length) % $items.length;
    $items.removeClass('active').eq(active).addClass('active');
    var el = $items.eq(active)[0];
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' });
  }

  function choose(i) {
  var s = shown[i];
  if (!s) return;

  Shiny.setInputValue('search_selection', {
    value: s.t,
    type: s.k
  }, {priority: 'event'});

  $inp().val(s.t);
  hideSuggest();
  $inp().focus();
}

  Shiny.addCustomMessageHandler('search_suggest', function(m) {
    var l = [].concat(m.labels || []);
    var k = [].concat(m.types || []);
    SUG = l.map(function(x, i) { return { t: String(x), k: k[i] }; });
  });

  $inp().attr({
    autocomplete: 'off',
    spellcheck: 'false',
    role: 'combobox',
    'aria-autocomplete': 'list',
    'aria-controls': 'sb_suggest',
    'aria-expanded': 'false'
  });
  syncHasText();

  $(document).on('input focusin', '#global_search', function() {
    syncHasText();
    renderSuggest();
  });

  $(document).on('focusout', '#global_search', function() {
    setTimeout(hideSuggest, 120);
  });

  $(document).on('shiny:inputchanged', function(e) {
    if (e.name === 'global_search') {
      $('.sb-search').toggleClass('has-text', !!e.value && String(e.value).length > 0);
    }
  });

  $(document).on('keydown', '#global_search', function(e) {
    var open = $box().hasClass('open');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!open) renderSuggest();
      setActive(active + 1);
    } else if (e.key === 'ArrowUp') {
      if (open) {
        e.preventDefault();
        setActive(active < 0 ? shown.length - 1 : active - 1);
      }
    } else if (e.key === 'Enter') {
      if (open && active >= 0) { e.preventDefault(); choose(active); }
      else hideSuggest();
    } else if (e.key === 'Escape') {
      if (open) hideSuggest();
      else if (this.value) $(this).val('').trigger('input');
      else this.blur();
    }
  });

  $(document).on('mousedown', '.sb-sug-item', function(e) {
    e.preventDefault();
    choose(parseInt($(this).attr('data-i'), 10));
  });

  $(document).on('mousemove', '.sb-sug-item', function() {
    var n = parseInt($(this).attr('data-i'), 10);
    if (n !== active) setActive(n);
  });

  $(document).on('click', '#sb_search_clear', function(e) {
    e.preventDefault();
    $inp().val('').trigger('input');
    hideSuggest();
    $inp().focus();
  });

  $(document).on('keydown', function(e) {
    var t = e.target || {};
    var tag = (t.tagName || '').toLowerCase();
    var typing = tag === 'input' || tag === 'textarea' || tag === 'select' ||
                 t.isContentEditable;
    var isK = (e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'k';
    if ((e.key === '/' && !typing) || isK) {
      if ($inp().is(':visible')) {
        e.preventDefault();
        $inp().focus().select();
      }
    }
  });

  // ==========================================================
  // TABLE NOTES: the (i) button only tells the server which row
  // was clicked. The server answers with a modal, so nothing is
  // shown until the button is pressed.
  // ==========================================================
  $(document).on('click', 'button.note-btn', function(e) {
    e.preventDefault();
    e.stopPropagation();
    var b = this;
    Shiny.setInputValue('note_click', {
      title: b.getAttribute('data-title') || '',
      sub:   b.getAttribute('data-sub') || '',
      note:  b.getAttribute('data-note') || ''
    }, { priority: 'event' });
  });
});
)---"

# Cluster bubble icon for the Data Centers map.
#   orange -> any site inside is highlighted (coming online soon)
#   heat   -> colour of the largest current site inside
#   gray   -> everything inside has no current capacity
CLUSTER_ICON_JS <- r"---(
function(cluster) {
  var kids = cluster.getAllChildMarkers();
  var best = null, anyOrange = false, anyGray = false;
  kids.forEach(function(m) {
    var o = m.options;
    var fc = (o.fillColor || '').toUpperCase();
    if (fc === '#F59E0B') anyOrange = true;
    if (fc === '#9CA3AF') anyGray = true;
    if (fc !== '#9CA3AF' && fc !== '#FFFFFF' &&
        (best === null || o.radius > best.options.radius)) best = m;
  });
  var col = anyOrange ? '#F59E0B'
          : (best ? best.options.fillColor
          : (anyGray ? '#9CA3AF' : '#FFFFFF'));
  var isWhite = (String(col).toUpperCase() === '#FFFFFF');
  var n = cluster.getChildCount();
  return L.divIcon({
    html: '<div style="background:' + col + ';' +
          'color:' + (isWhite ? '#111827' : '#fff') + ';' +
          'text-shadow:' + (isWhite ? 'none' : '0 0 3px #000') + ';' +
          'font-weight:700;width:36px;height:36px;line-height:36px;' +
          'border-radius:50%;text-align:center;' +
          'border:2px solid ' + (isWhite ? '#6B7280' : 'rgba(255,255,255,.6)') + ';' +
          'opacity:.9">' + n + '</div>',
    className: '',
    iconSize: L.point(36, 36)
  });
}
)---"

# ============================================================
# CSS
# Plain CSS in its own <style> tag (not run through Sass), so a
# rule can never be silently dropped by the theme compiler.
# ============================================================

APP_CSS <- r"---(
:root {
  --bg: #080C12;
  --panel: #0F1620;
  --panel-2: #131C28;
  --line: #1E2A3A;
  --line-2: #2A3A50;
  --ink: #E6EDF6;
  --ink-2: #A3B2C6;
  --ink-3: #6B7C93;
  --blue: #3B82F6;
  --blue-2: #60A5FA;
  --amber: #F59E0B;
  --green: #22C55E;
}

html, body, .bslib-page-sidebar, .bslib-sidebar-layout,
.bslib-sidebar-layout > .main, .tab-content, .container-fluid {
  background-color: var(--bg) !important;
  color: var(--ink) !important;
}

body {
  background-image:
    radial-gradient(1100px 500px at 85% -10%, rgba(59,130,246,.10), transparent 60%),
    radial-gradient(800px 400px at -10% 110%, rgba(59,130,246,.06), transparent 60%);
  background-attachment: fixed;
}

/* ---------- Header ---------- */
.bslib-page-sidebar > header, .bslib-page-title {
  background: transparent !important;
}
.app-title {
  display: flex; align-items: center; gap: 14px; width: 100%;
}
.app-title img { height: 80px; border-radius: 8px; }
.app-title-text { line-height: 1.15; }
.app-title-name { font-size: 20px; font-weight: 700; letter-spacing: -.01em; color: #F8FAFC; }
.app-title-sub  { font-size: 12.5px; color: var(--ink-3); margin-top: 2px; }
.app-title-right { margin-left: auto; display: flex; align-items: center; gap: 10px; }

.pill {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 5px 12px; border-radius: 999px;
  border: 1px solid var(--line-2); background: rgba(15,22,32,.8);
  font-size: 12px; color: var(--ink-2); text-decoration: none;
}
.pill .live-dot {
  width: 7px; height: 7px; border-radius: 50%; background: var(--green);
  box-shadow: 0 0 0 3px rgba(34,197,94,.18);
}
a.pill:hover { border-color: var(--blue); color: #fff; }
.pill-warn {
  border-color: rgba(245,158,11,.5); color: #FCD34D;
  background: rgba(245,158,11,.10); cursor: pointer;
}
a.pill.pill-warn:hover { border-color: var(--amber); color: #FDE68A; }

/* ---------- Tabs ---------- */
.nav-tabs { border-bottom: 1px solid var(--line) !important; margin-bottom: 18px; gap: 4px; }
.nav-tabs .nav-link {
  color: var(--ink-3) !important; border: 0 !important; background: transparent !important;
  padding: 10px 16px; font-weight: 600; font-size: 14px;
  border-bottom: 2px solid transparent !important;
}
.nav-tabs .nav-link:hover { color: var(--ink) !important; }
.nav-tabs .nav-link.active {
  color: #fff !important; border-bottom-color: var(--blue) !important;
}

/* ---------- Cards ---------- */
.card, .card-body {
  background-color: var(--panel) !important;
  color: var(--ink) !important;
  border-color: var(--line) !important;
}
.card {
  border-radius: 14px !important;
  box-shadow: 0 1px 0 rgba(255,255,255,.03) inset, 0 12px 28px rgba(0,0,0,.28);
  margin-bottom: 16px;
}
.card-header, .card-footer {
  background-color: var(--panel) !important;
  color: var(--ink) !important;
  border-color: var(--line) !important;
}
.dc-card-head { display: flex; align-items: center; gap: 10px; padding: 12px 18px !important; }
.dc-card-ic {
  width: 28px; height: 28px; border-radius: 8px;
  display: inline-flex; align-items: center; justify-content: center;
  background: rgba(59,130,246,.14); color: var(--blue-2); font-size: 13px;
}
.dc-card-title { font-weight: 650; font-size: 14.5px; color: #F1F5F9; }
.dc-card-hint { margin-left: auto; font-size: 12px; color: var(--ink-3); }

/* ---------- KPI tiles ---------- */
.kpi {
  display: flex; align-items: center; gap: 16px;
  padding: 18px 20px; border-radius: 14px;
  background: var(--panel); border: 1px solid var(--line);
  position: relative; overflow: hidden;
}
.kpi::before {
  content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 3px;
  background: var(--accent, var(--blue));
}
.kpi-icon {
  flex: none; width: 46px; height: 46px; border-radius: 12px;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 19px; color: var(--accent, var(--blue));
  background: color-mix(in srgb, var(--accent, #3B82F6) 14%, transparent);
}
.kpi-label { font-size: 13px; color: var(--ink-2); }
.kpi-value {
  font-size: 30px; font-weight: 750; letter-spacing: -.02em;
  color: #F8FAFC; line-height: 1.15; font-variant-numeric: tabular-nums;
}
.kpi-yellow { --accent: #FACC15; }
.kpi-blue   { --accent: #3B82F6; }
.kpi-teal   { --accent: #2DD4BF; }
.kpi-violet  { --accent: #A855F7; }

/* ---------- Sidebar ---------- */
.bslib-sidebar-layout > .sidebar {
  background-color: #0B1119 !important;
  color: var(--ink) !important;
  border-right: 1px solid var(--line) !important;
}
.sb-summary {
  background: var(--panel-2); border: 1px solid var(--line);
  border-radius: 12px; padding: 10px 14px; margin-bottom: 12px;
}
.sb-summary-main { font-size: 13px; color: var(--ink-2); }
.sb-summary-main b { font-size: 22px; color: var(--blue-2); }

.form-control, .selectize-input {
  background-color: #080C12 !important; color: var(--ink) !important;
  border-color: var(--line) !important;
}
.selectize-dropdown {
  background-color: var(--panel) !important; color: var(--ink) !important;
  border-color: var(--line) !important;
}
.form-control:focus, .selectize-input.focus {
  box-shadow: 0 0 0 .2rem rgba(59,130,246,.25) !important;
  border-color: var(--blue) !important;
}

/* Search bar */
.sb-search { position: relative; margin-bottom: 12px; }
.sb-search .form-group, .sb-search .shiny-input-container { margin: 0 !important; width: 100% !important; }
.sb-search input#global_search {
  height: 42px; padding: 0 44px 0 40px; border-radius: 999px !important;
  background-color: #080C12 !important; border: 1px solid var(--line-2) !important;
  color: var(--ink) !important; font-size: 13.5px;
  transition: border-color .15s ease, box-shadow .15s ease;
}
.sb-search input#global_search::placeholder { color: var(--ink-3); }
.sb-search input#global_search:focus {
  border-color: var(--blue) !important;
  box-shadow: 0 0 0 .22rem rgba(59,130,246,.22) !important; outline: none;
}
.sb-search-icon {
  position: absolute; left: 15px; top: 50%; transform: translateY(-50%);
  color: var(--ink-3); font-size: 13px; pointer-events: none; z-index: 3;
}
.sb-search:focus-within .sb-search-icon { color: var(--blue-2); }
.sb-search-kbd {
  position: absolute; right: 12px; top: 50%; transform: translateY(-50%);
  min-width: 22px; height: 22px; padding: 0 6px; line-height: 20px; text-align: center;
  border: 1px solid #334155; border-bottom-width: 2px; border-radius: 6px;
  background: var(--panel); color: var(--ink-2);
  font-family: inherit; font-size: 11px; font-weight: 600;
  pointer-events: none; z-index: 3;
}
.sb-search:focus-within .sb-search-kbd, .sb-search.has-text .sb-search-kbd { display: none; }
@media (hover: none) { .sb-search-kbd { display: none; } }
.sb-search-clear {
  position: absolute; right: 10px; top: 50%; transform: translateY(-50%);
  width: 22px; height: 22px; padding: 0; border: 0; line-height: 21px; text-align: center;
  border-radius: 50%; background: #1E293B; color: var(--ink-2);
  font-size: 15px; font-weight: 700; display: none; z-index: 3; cursor: pointer;
}
.sb-search-clear:hover { background: #334155; color: #F87171; }
.sb-search.has-text .sb-search-clear { display: block; }
.sb-suggest {
  display: none; position: absolute; left: 0; right: 0; top: calc(100% + 6px);
  z-index: 1050; max-height: 320px; overflow-y: auto; padding: 5px;
  background: var(--panel); border: 1px solid var(--line-2); border-radius: 12px;
  box-shadow: 0 14px 32px rgba(0,0,0,.55);
}
.sb-suggest.open { display: block; }
.sb-sug-item {
  display: flex; align-items: center; justify-content: space-between;
  gap: 10px; padding: 7px 10px; border-radius: 8px;
  font-size: 13px; color: #CBD5E1; cursor: pointer;
}
.sb-sug-item.active { background: rgba(59,130,246,.18); color: #F1F5F9; }
.sb-sug-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sb-sug-text mark { background: transparent; color: var(--blue-2); font-weight: 700; padding: 0; }
.sb-sug-type {
  flex: none; padding: 1px 8px; font-size: 11px; color: var(--ink-3);
  border: 1px solid var(--line); border-radius: 999px;
}
.sb-sug-item.active .sb-sug-type { border-color: var(--blue); color: #93C5FD; }

.sb-search-status {
  display: flex; align-items: flex-start; gap: 8px; margin: -4px 4px 12px;
  font-size: 12px; line-height: 1.4; color: var(--ink-2);
}
.sb-search-status b { color: var(--ink); font-weight: 600; }
.sb-search-status .sb-dot {
  flex: none; width: 7px; height: 7px; margin-top: 5px; border-radius: 50%;
  background: var(--green); box-shadow: 0 0 6px rgba(34,197,94,.6);
}
.sb-search-status.is-empty { color: #FBBF24; }
.sb-search-status.is-empty b { color: #FDE68A; }

.sb-presets { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 12px; }
.sb-presets .btn { flex: 1; }
.sb-presets .btn.is-active {
  background: var(--blue); border-color: var(--blue); color: #fff;
}
.sb-note {
  display: flex; gap: 8px; align-items: flex-start;
  font-size: 12px; line-height: 1.4; color: var(--ink-3); margin: -2px 2px 10px;
}
.sb-chips { margin-bottom: 8px; }
.sb-chip {
  display: inline-flex; align-items: center; gap: 6px;
  background: #1E293B; border: 1px solid #334155; border-radius: 999px;
  padding: 2px 10px; font-size: 12px; margin: 0 6px 6px 0;
}
.sb-chip-x { color: var(--ink-2); text-decoration: none; font-weight: 700; }
.sb-chip-x:hover { color: #F87171; }
.sb-footer { font-size: 11px; color: var(--ink-3); text-align: center; margin-top: 10px; }

.accordion {
  --bs-accordion-bg: transparent;
  --bs-accordion-border-color: var(--line);
  --bs-accordion-btn-color: var(--ink);
  --bs-accordion-btn-bg: var(--panel);
  --bs-accordion-active-bg: #0B1220;
  --bs-accordion-active-color: #93C5FD;
  --bs-accordion-btn-focus-box-shadow: 0 0 0 .2rem rgba(59,130,246,.25);
}
.accordion-button { font-size: 13px; font-weight: 600; }
.accordion-button::after { filter: invert(1); }

.seg-toggle .shiny-options-group {
  display: flex; background: #080C12; border: 1px solid var(--line);
  border-radius: 8px; padding: 3px;
}
.seg-toggle .form-check { flex: 1; margin: 0; padding: 0; }
.seg-toggle .form-check-input { position: absolute; opacity: 0; }
.seg-toggle .form-check-label {
  display: block; text-align: center; padding: 6px 0; border-radius: 6px;
  cursor: pointer; color: var(--ink-2); font-size: 13px; transition: all .15s;
}
.seg-toggle .form-check-input:checked + .form-check-label { background: var(--blue); color: #fff; }

.sb-section-label { font-size: 12px; font-weight: 600; color: var(--ink-2); margin: 14px 0 6px; }
.sb-section-label:first-child { margin-top: 2px; }

.range-wrap .irs-from, .range-wrap .irs-to, .range-wrap .irs-single { display: none !important; }
.range-inputs { display: flex; align-items: center; gap: 8px; margin-bottom: 2px; }
.range-inputs .form-group, .range-inputs .shiny-input-container {
  margin: 0 !important; flex: 1; width: auto !important;
}
.range-inputs input { text-align: center; font-weight: 600; color: var(--blue-2) !important; }
.range-dash { color: var(--ink-3); }

/* ---------- Data tables ---------- */
.dataTables_wrapper { background-color: var(--panel) !important; color: var(--ink) !important; }
table.dataTable, table.dataTable td, table.dataTable th {
  background-color: var(--panel) !important; color: var(--ink) !important;
  border-color: var(--line) !important;
}
table.dataTable thead th {
  color: var(--ink-2) !important; font-weight: 600; font-size: 12.5px;
  border-bottom: 1px solid var(--line-2) !important;
}
table.dataTable { width: 100% !important; }
.dataTables_wrapper { width: 100%; }
table.dataTable tbody td { font-size: 13px; vertical-align: middle; }
table.dataTable tbody tr:hover td { background-color: #17212F !important; }
.dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter,
.dataTables_wrapper .dataTables_info, .dataTables_wrapper .dataTables_paginate { color: var(--ink-2) !important; }
.dataTables_wrapper .form-control, .dataTables_wrapper select {
  background-color: #080C12 !important; color: var(--ink) !important; border-color: var(--line) !important;
}
.page-link { background: var(--panel) !important; color: var(--ink-2) !important; border-color: var(--line) !important; }
.page-item.active .page-link { background: var(--blue) !important; color: #fff !important; border-color: var(--blue) !important; }

/* ---------- Notes: table button ---------- */
.note-btn {
  width: 26px; height: 26px; padding: 0; border-radius: 50%;
  border: 1px solid #334155; background: #0B1220; color: #93C5FD;
  font: italic 700 13px/24px Georgia, serif; cursor: pointer;
  transition: background .15s ease, border-color .15s ease, box-shadow .15s ease;
}
.note-btn:hover, .note-btn:focus-visible {
  background: var(--blue); border-color: var(--blue); color: #fff;
  box-shadow: 0 0 0 .22rem rgba(59,130,246,.25); outline: none;
}
.note-none { color: #334155; }

/* ---------- Notes: modal (opened from the table) ---------- */
.modal-content {
  background: var(--panel) !important; color: var(--ink) !important;
  border: 1px solid var(--line-2) !important; border-radius: 16px !important;
  box-shadow: 0 24px 60px rgba(0,0,0,.65);
}
.modal-header, .modal-footer { border-color: var(--line) !important; }
.modal-header { background: linear-gradient(135deg, rgba(59,130,246,.18), rgba(59,130,246,0)); }
.note-modal-op { font-size: 17px; font-weight: 700; color: #F8FAFC; }
.note-modal-sub { font-size: 13px; color: var(--ink-2); margin-top: 2px; }
.note-modal-body { font-size: 14px; line-height: 1.65; color: #D5DEEA; white-space: pre-wrap; max-height: 55vh; overflow-y: auto; }
.modal-backdrop.show { opacity: .65; }

/* ---------- Notes: map popup (native <details>, closed by default) ---------- */
.map-note { margin-top: 8px; }
.map-note > summary {
  list-style: none; cursor: pointer;
  display: inline-flex; align-items: center; gap: 6px;
  padding: 3px 10px 3px 4px; border: 1px solid #BFDBFE; border-radius: 999px;
  background: #EFF6FF; color: #1D4ED8; font-size: 12px; font-weight: 600;
}
.map-note > summary::-webkit-details-marker { display: none; }
.map-note > summary::marker { content: ""; }
.map-note > summary:hover { background: #DBEAFE; }
.map-note-i {
  width: 18px; height: 18px; border-radius: 50%; background: #3B82F6; color: #fff;
  font: italic 700 11px/18px Georgia, serif; text-align: center;
}
.map-note-body {
  margin-top: 8px; padding: 8px 10px; background: #F8FAFC;
  border-left: 3px solid #3B82F6; border-radius: 6px;
  font-size: 12px; line-height: 1.5; color: #374151;
  max-height: 160px; overflow-y: auto; white-space: pre-wrap;
}

/* ---------- Map ---------- */
.leaflet-container { background-color: var(--bg) !important; border-radius: 0 0 14px 14px; }
.dc-legend {
  background: rgba(255,255,255,.95); color: #1F2937; padding: 8px 12px;
  border-radius: 8px; box-shadow: 0 1px 5px rgba(0,0,0,.35);
  font-size: 12px; line-height: 1.3; min-width: 150px;
}
.dc-legend-title { font-weight: 700; margin-bottom: 5px; }
.dc-legend-bar { height: 10px; border-radius: 5px; background: linear-gradient(to right, #DBEAFE, #60A5FA, #1D4ED8, #0A1A4A); }
.dc-legend-scale { display: flex; justify-content: space-between; font-size: 11px; color: #4B5563; margin-top: 2px; }
.dc-legend-row { display: flex; align-items: center; gap: 7px; margin-top: 6px; }
.dc-legend-dot { width: 11px; height: 11px; border-radius: 50%; display: inline-block; flex: none; }

/* ---------- Map / Satellite switch (Leaflet layers control) ----------
   Restyles Leaflet's built-in base-layer radio list as a frosted,
   pill-shaped segmented control that matches the app's dark theme. */
.leaflet-control-layers.leaflet-control {
  background: rgba(8,12,18,.82) !important;
  -webkit-backdrop-filter: blur(10px); backdrop-filter: blur(10px);
  border: 1px solid rgba(148,163,184,.28) !important;
  border-radius: 999px !important;
  box-shadow: 0 8px 22px rgba(0,0,0,.35), 0 1px 0 rgba(255,255,255,.06) inset !important;
  padding: 3px !important;
  color: var(--ink) !important;
  font-family: 'Inter', system-ui, sans-serif;
}
.leaflet-control-layers-expanded { padding: 3px !important; }
.leaflet-control-layers-list { margin: 0 !important; }
.leaflet-control-layers-base { display: flex; gap: 2px; }
.leaflet-control-layers-separator,
.leaflet-control-layers-overlays { display: none !important; }
.leaflet-control-layers-base label {
  position: relative; margin: 0 !important; cursor: pointer;
  border-radius: 999px; padding: 6px 14px;
  font-size: 12px; font-weight: 650; letter-spacing: .01em; line-height: 1.2;
  color: var(--ink-2);
  transition: background .18s ease, color .18s ease, box-shadow .18s ease;
}
.leaflet-control-layers-base label > span,
.leaflet-control-layers-base label > div { display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
.leaflet-control-layers-base label:hover { color: #F8FAFC; background: rgba(148,163,184,.14); }
/* hide the native radio but keep it focusable for keyboard users */
.leaflet-control-layers-base input.leaflet-control-layers-selector {
  position: absolute; opacity: 0; width: 1px; height: 1px; margin: 0; pointer-events: none;
}
/* small icon dot before each option */
.leaflet-control-layers-base label > span::before,
.leaflet-control-layers-base label > div::before {
  content: ""; width: 8px; height: 8px; border-radius: 50%;
  background: currentColor; opacity: .55; flex: none;
}
/* selected option */
.leaflet-control-layers-base label:has(input:checked) {
  background: linear-gradient(135deg, #3B82F6, #2563EB);
  color: #fff;
  box-shadow: 0 2px 10px rgba(59,130,246,.45);
}
.leaflet-control-layers-base label:has(input:checked) > span::before,
.leaflet-control-layers-base label:has(input:checked) > div::before {
  opacity: 1; background: #fff; box-shadow: 0 0 0 3px rgba(255,255,255,.25);
}
/* fallback for browsers without :has() */
.leaflet-control-layers-base input:checked + span { color: #fff; font-weight: 750; }
.leaflet-control-layers-base label:has(input:focus-visible) {
  outline: 2px solid var(--blue-2); outline-offset: 1px;
}

/* ---------- Misc ---------- */
.btn-outline-primary { --bs-btn-color: #93C5FD; --bs-btn-border-color: #2B4A7A; --bs-btn-hover-bg: var(--blue); --bs-btn-hover-border-color: var(--blue); }
.version-note { color: var(--ink-2); font-size: 13.5px; max-width: 70ch; }
.dq-ok { color: var(--ink-2); }
.dq-ok svg, .dq-ok .fa { color: var(--green); margin-right: 6px; }
:focus-visible { outline: 2px solid var(--blue-2); outline-offset: 2px; }
@media (max-width: 768px) {
  .app-title-sub { display: none; }
  .kpi-value { font-size: 24px; }
  .leaflet-control-layers-base label { padding: 5px 11px; font-size: 11.5px; }
}
@media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }

/* ---------- Fancy "Coming soon" toggle ---------- */
.btn.hl-btn {
  position: relative; overflow: hidden;
  display: flex; align-items: center; gap: 12px;
  padding: 11px 14px; text-align: left;
  border-radius: 14px; border: 1px solid #3A2F14;
  background: linear-gradient(135deg, #1A1508, #0F1620 70%);
  color: var(--ink);
  transition: transform .15s ease, box-shadow .25s ease, border-color .25s ease, background .25s ease;
}
.btn.hl-btn:hover {
  transform: translateY(-1px); border-color: var(--amber);
  box-shadow: 0 6px 18px rgba(245,158,11,.18);
}
.btn.hl-btn:active { transform: translateY(0) scale(.99); }

.hl-icon {
  flex: none; width: 36px; height: 36px; border-radius: 10px;
  display: inline-flex; align-items: center; justify-content: center;
  color: var(--amber); background: rgba(245,158,11,.14); font-size: 15px;
  transition: all .25s ease;
}
.hl-text { display: flex; flex-direction: column; line-height: 1.2; min-width: 0; }
.hl-title { font-weight: 700; font-size: 13.5px; color: #F8FAFC; }
.hl-sub   { font-size: 11.5px; color: var(--ink-3); transition: color .25s ease; }

.hl-count {
  margin-left: auto; min-width: 26px; padding: 2px 8px; text-align: center;
  border-radius: 999px; font-size: 12px; font-weight: 700;
  color: #FCD34D; background: rgba(245,158,11,.14);
  border: 1px solid rgba(245,158,11,.35); font-variant-numeric: tabular-nums;
}

/* ON state */
.btn.hl-btn.is-on {
  border-color: var(--amber);
  background: linear-gradient(135deg, #3A2606, #1F1608 70%);
  box-shadow: 0 0 0 1px rgba(245,158,11,.5), 0 0 22px rgba(245,158,11,.28);
  animation: hlGlow 2.4s ease-in-out infinite;
}
.btn.hl-btn.is-on::after {          /* light sweep */
  content: ""; position: absolute; top: 0; bottom: 0; left: -60%; width: 40%;
  background: linear-gradient(100deg, transparent, rgba(255,255,255,.14), transparent);
  transform: skewX(-20deg); animation: hlSweep 3s ease-in-out infinite;
  pointer-events: none;
}
.is-on .hl-icon { background: var(--amber); color: #1A1204; animation: hlPulse 1.6s ease-out infinite; }
.is-on .hl-sub  { color: #FCD34D; }

@keyframes hlPulse {
  0%   { box-shadow: 0 0 0 0 rgba(245,158,11,.55); }
  100% { box-shadow: 0 0 0 12px rgba(245,158,11,0); }
}
@keyframes hlGlow {
  0%,100% { box-shadow: 0 0 0 1px rgba(245,158,11,.5), 0 0 16px rgba(245,158,11,.22); }
  50%     { box-shadow: 0 0 0 1px rgba(245,158,11,.7), 0 0 28px rgba(245,158,11,.40); }
}
@keyframes hlSweep {
  0%   { left: -60%; }
  60%,100% { left: 130%; }
}
.hl-count::before { content: ""; }
.is-on .hl-count {
  background: var(--amber); color: #1A1204; border-color: var(--amber);
}
.is-on .hl-title::after {
  content: " \2713"; color: #FCD34D;
}

/* ---------- Donut toggle (Country / City) ---------- */
.seg-mini { margin-left: 12px; width: 150px; flex: none; }
.seg-mini .shiny-input-container, .seg-mini .form-group { margin: 0 !important; width: 100% !important; }
.seg-mini .form-check-label { padding: 4px 0; font-size: 12px; }
)---"

# ---------- Find Capacity tab ----------
FIND_CSS <- r"---(
.fc-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:16px; margin-bottom:16px; }
.fc-card { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:16px 18px; position:relative; overflow:hidden; }
.fc-card.rank-1 { border-color:var(--blue); box-shadow:0 0 0 1px rgba(59,130,246,.4), 0 12px 28px rgba(0,0,0,.35); }
.fc-rank { font-size:11px; font-weight:700; letter-spacing:.08em; color:var(--blue-2); }
.fc-name { font-size:16px; font-weight:700; color:#F8FAFC; margin-top:2px; }
.fc-sub  { font-size:12.5px; color:var(--ink-3); }
.fc-score { display:flex; align-items:center; gap:10px; margin:12px 0 8px; }
.fc-bar { flex:1; height:8px; border-radius:99px; background:#1E293B; overflow:hidden; }
.fc-bar > span { display:block; height:100%; border-radius:99px; background:linear-gradient(90deg,#60A5FA,#3B82F6); }
.fc-pct { font-weight:750; font-size:20px; color:#F8FAFC; font-variant-numeric:tabular-nums; }
.fc-why { font-size:12.5px; color:var(--ink-2); line-height:1.5; }
.fc-gap { color:#FBBF24; }
.fc-tag { display:inline-block; font-size:11px; padding:1px 8px; border-radius:99px; border:1px solid var(--line-2); color:var(--ink-2); margin-right:6px; margin-bottom:4px; }
.fc-firm { border-color:rgba(34,197,94,.45); color:#86EFAC; }
.fc-summary { padding:14px 18px; border-radius:12px; background:var(--panel-2); border:1px solid var(--line); margin-bottom:16px; font-size:14px; }
.fc-summary b { color:var(--blue-2); }
)---"

# ---------- Highlight window timeline ----------
TIMELINE_CSS <- r"---(
.hl-tl {
  margin-top: 10px; padding: 12px 12px 10px; border-radius: 14px;
  border: 1px solid var(--line); background: var(--panel-2); opacity: .62;
  transition: opacity .25s ease, border-color .25s ease, box-shadow .25s ease;
}
.hl-tl.is-on {
  opacity: 1; border-color: rgba(245,158,11,.45);
  box-shadow: 0 0 0 1px rgba(245,158,11,.18), 0 8px 22px rgba(245,158,11,.10);
}
.tl-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
.tl-title { font-size: 11px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--ink-3); }
.tl-chips { display: flex; gap: 4px; }
.tl-chip {
  border: 1px solid var(--line-2); background: transparent; color: var(--ink-2);
  font-size: 11px; font-weight: 600; padding: 1px 8px; border-radius: 99px; cursor: pointer;
  transition: border-color .15s ease, color .15s ease;
}
.tl-chip:hover { border-color: var(--amber); color: #FCD34D; }
.tl-strip { display: flex; gap: 3px; align-items: stretch; height: 92px; user-select: none; -webkit-user-select: none; }
.tl-col {
  flex: 1 1 0; min-width: 0; padding: 0; border: 0; background: transparent;
  display: flex; flex-direction: column; align-items: stretch; cursor: pointer; color: var(--ink-3);
}
.tl-bar { flex: 1; display: flex; align-items: flex-end; border-radius: 5px; background: rgba(148,163,184,.07); overflow: hidden; }
.tl-fill { width: 100%; border-radius: 4px 4px 0 0; background: #243247; transition: background .2s ease, filter .2s ease; }
.tl-col.in .tl-bar { background: rgba(59,130,246,.10); }
.tl-col.in .tl-fill { background: linear-gradient(180deg, #93C5FD, #3B82F6); }
.hl-tl.is-on .tl-col.in .tl-bar { background: rgba(245,158,11,.12); }
.hl-tl.is-on .tl-col.in .tl-fill { background: linear-gradient(180deg, #FCD34D, #F59E0B); }
.tl-col:hover .tl-fill { filter: brightness(1.2); }
.tl-col.edge .tl-fill { filter: brightness(1.3); }
.tl-col.edge .tl-q { color: #F8FAFC; }
.tl-col:focus-visible { outline: 2px solid var(--blue-2); outline-offset: 1px; border-radius: 5px; }
.tl-q { font-size: 9.5px; font-weight: 700; margin-top: 4px; line-height: 1; text-align: center; }
.tl-y { font-size: 9px; line-height: 1.4; height: 13px; color: var(--ink-3); text-align: center; }
.tl-read { margin-top: 6px; font-size: 12px; color: var(--ink-2); text-align: center; }
.tl-read b { color: #F8FAFC; }
.hl-tl.is-on .tl-read b { color: #FCD34D; }
)---"

# ============================================================
# UI
# ============================================================

ui <- page_sidebar(
  title = div(
    class = "app-title",
    tags$img(src = "logo.png", alt = "",
             onerror = "this.style.display='none'"),
    div(
      class = "app-title-text",
      div(class = "app-title-name", "Data Center Capacity"),
      div(class = "app-title-sub", "Current sites and upcoming capacity across operators")
    ),
    div(
      class = "app-title-right",
      if (nzchar(SHAREPOINT_URL)) {
        tags$a(class = "pill", href = SHAREPOINT_URL, target = "_blank", rel = "noopener",
               icon("file-excel"), "Master workbook")
      },
      uiOutput("hdr_flags", inline = TRUE),
      uiOutput("hdr_fresh", inline = TRUE)
    )
  ),
  
  theme = bs_theme(
    version = 5,
    bg = "#080C12",
    fg = "#E6EDF6",
    primary = "#3B82F6",
    secondary = "#64748B",
    success = "#22C55E",
    warning = "#F59E0B",
    danger = "#EF4444",
    base_font = font_google("Inter", local = TRUE),
    heading_font = font_google("Inter", local = TRUE)
  ),
  
  tags$head(
    tags$style(HTML(APP_CSS)),
    tags$style(HTML(FIND_CSS)),
    tags$style(HTML(TIMELINE_CSS)),
    tags$script(HTML(APP_JS))
  ),
  
  # ==========================================================
  # SIDEBAR
  # ==========================================================
  
  sidebar = sidebar(
    width = 320,
    open = "desktop",
    
    uiOutput("sb_summary"),
    
    # ---------------- Search ----------------
    conditionalPanel(
      condition = "input.main_tabs == 'Data Centers' || input.main_tabs == 'Upcoming Capacity'",
      
      div(
        class = "sb-search",
        tags$span(class = "sb-search-icon", icon("magnifying-glass")),
        textInput(
          "global_search", NULL, width = "100%",
          placeholder = "Search operator, city, state, country"
        ),
        tags$kbd(class = "sb-search-kbd", "/"),
        tags$button(
          type = "button", id = "sb_search_clear", class = "sb-search-clear",
          `aria-label` = "Clear search", title = "Clear search (Esc)",
          HTML("&times;")
        ),
        div(id = "sb_suggest", class = "sb-suggest", role = "listbox")
      ),
      
      uiOutput("search_status")
    ),
    
    # ---------------- Data Centers filters ----------------
    conditionalPanel(
      condition = "input.main_tabs == 'Data Centers'",
      
      uiOutput("filter_chips"),
      
      accordion(
        id = "sb_acc",
        open = c("loc", "cap"),
        multiple = TRUE,
        
        accordion_panel(
          "Location", value = "loc", icon = icon("location-dot"),
          div(class = "seg-toggle",
              radioButtons("scope", NULL,
                           choices = c("US only" = "us", "Global" = "global"),
                           selected = "global", inline = TRUE)),
          selectizeInput("state_filter", "State (US)", choices = NULL, multiple = TRUE,
                         options = list(placeholder = "All states",
                                        plugins = list("remove_button"))),
          conditionalPanel(
            condition = "input.scope == 'global'",
            selectizeInput("country_filter", "Country (non-US)", choices = NULL, multiple = TRUE,
                           options = list(placeholder = "All countries",
                                          plugins = list("remove_button")))
          ),
          selectizeInput("city_filter", "City", choices = NULL, multiple = TRUE,
                         options = list(placeholder = "All cities",
                                        plugins = list("remove_button")))
        ),
        
        accordion_panel(
          "Operator", value = "op", icon = icon("building"),
          selectizeInput("operator_filter", NULL, choices = NULL, multiple = TRUE,
                         options = list(placeholder = "All operators",
                                        plugins = list("remove_button")))
        ),
        
        accordion_panel(
          "Capacity (MW)", value = "cap", icon = icon("bolt"),
          range_slider("capacity_filter", 0, 100, c(0, 100), step = 1)
        )
      ),
      
      actionButton(
        "highlight_toggle",
        label = tagList(
          span(class = "hl-icon", icon("bolt")),
          span(class = "hl-text",
               span(class = "hl-title", textOutput("hl_count", inline = TRUE)),
               span(class = "hl-sub", "Click to highlight on map"))
        ),
        class = "hl-btn w-100 mt-3",
        `aria-pressed` = "false",
        title = "Highlight sites with upcoming capacity inside the chosen window"
      ),
      
      uiOutput("hl_timeline")
    ),
    
    # ---------------- Upcoming Capacity filters ----------------
    conditionalPanel(
      condition = "input.main_tabs == 'Upcoming Capacity'",
      
      div(class = "sb-section-label", "Scope"),
      div(class = "seg-toggle",
          radioButtons("pipeline_scope", NULL,
                       choices = c("US only" = "us", "Global" = "global"),
                       selected = "global", inline = TRUE)),
      
      conditionalPanel(
        condition = "input.pipeline_scope == 'us'",
        div(class = "sb-section-label", "State"),
        selectizeInput("pipeline_state_filter", NULL, choices = NULL, multiple = TRUE,
                       options = list(placeholder = "All states",
                                      plugins = list("remove_button")))
      ),
      
      conditionalPanel(
        condition = "input.pipeline_scope == 'global'",
        div(class = "sb-section-label", "Country"),
        selectizeInput("pipeline_country_filter", NULL, choices = NULL, multiple = TRUE,
                       options = list(placeholder = "All countries",
                                      plugins = list("remove_button")))
      ),
      
      div(class = "sb-section-label", "Timing"),
      uiOutput("year_presets"),
      selectizeInput("quarter_filter", NULL, choices = NULL, multiple = TRUE,
                     options = list(placeholder = "All quarters",
                                    plugins = list("remove_button"))),
      checkboxInput("pipeline_include_past", "Include past quarters", value = FALSE),
      uiOutput("past_note"),
      
      div(class = "sb-section-label", "MW available"),
      range_slider("pipeline_mw_filter", 0, 100, c(0, 100), step = 0.1)
    ),
    
    hr(),
    actionButton("reset_filters", "Reset all filters",
                 icon = icon("rotate-left"), class = "w-100"),
    div(class = "sb-footer", textOutput("data_freshness"))
  ),
  
  # ==========================================================
  # TABS
  # ==========================================================
  
  navset_tab(
    id = "main_tabs",
    
    # --------------------------------------------------------
    # DATA CENTERS
    # --------------------------------------------------------
    nav_panel(
      "Data Centers",
      
      layout_columns(
        col_widths = c(4, 4, 4),
        kpi("yellow", "bolt", "Total capacity (MW est.)", "vb_capacity"),
        kpi("blue", "building", "Operators in view", "vb_operators"),
        kpi("teal", "city", "Cities in view", "vb_cities")
      ),
      
      card(
        full_screen = TRUE,
        card_title("earth-americas", "Map of total capacity",
                   "Bigger, darker dots = more MW. Hover or click a dot for details."),
        leafletOutput("map", height = 520)
      ),
      
      layout_columns(
        col_widths = c(7, 5),
        card(
          card_title("ranking-star", "Top operators", "By estimated MW in view"),
          plotlyOutput("top_operators_chart", height = "330px")
        ),
        card(
          card_header(
            class = "dc-card-head",
            span(class = "dc-card-ic", icon("globe")),
            span(class = "dc-card-title", textOutput("donut_title", inline = TRUE)),
            span(class = "dc-card-hint", "Share of MW in view"),
            div(class = "seg-toggle seg-mini",
                radioButtons("donut_by", NULL,
                             choices = c("Country" = "country", "City" = "city"),
                             selected = "country", inline = TRUE))
          ),
          plotlyOutput("share_chart", height = "330px")
        )
      ),
      
      card(
        card_title("table", "Matching rows", "Click the (i) button on a row to read its notes"),
        DTOutput("table"),
        card_footer(
          downloadButton(
            "download_filtered",
            "Download filtered results (CSV)",
            class = "btn-outline-primary btn-sm"
          )
        )
      )
    ),
    
    # --------------------------------------------------------
    # UPCOMING CAPACITY
    # --------------------------------------------------------
    nav_panel(
      "Upcoming Capacity",
      
      layout_columns(
        col_widths = c(4, 4, 4),
        kpi("yellow", "bolt", "Upcoming capacity (MW est.)", "vb_pipeline_mw"),
        kpi("violet", "clock", "Upcoming entries in view", "vb_pipeline_entries"),
        kpi("blue", "building", "Sites in view", "vb_pipeline_sites")
      ),
      
      card(
        full_screen = TRUE,
        card_title("map-location-dot", "Upcoming capacity map",
                   "Each dot is one site and quarter. Hover for the MW range."),
        leafletOutput("pipeline_map", height = 560)
      ),
      
      card(
        card_title("chart-column", "Upcoming capacity by quarter",
                   "Hover a bar to see the operator breakdown and MW range"),
        plotlyOutput("pipeline_quarter_chart", height = "380px")
      ),
      
      card(
        card_title("calendar-days", "Capacity coming available",
                   "In chronological order"),
        DTOutput("pipeline_table"),
        card_footer(
          downloadButton(
            "download_pipeline",
            "Download filtered results (CSV)",
            class = "btn-outline-primary btn-sm"
          ),
          downloadButton(
            "download_pipeline_pdf",
            "Download Client PDF",
            icon = icon("file-pdf"),
            class = "btn-outline-primary btn-sm ms-2"
          )
        )
      )
    ),
    
    # --------------------------------------------------------
    # FIND CAPACITY
    # --------------------------------------------------------
    nav_panel(
      "Find Capacity",
      
      card(
        card_title("bullseye", "What do you need?",
                   "Only sites with upcoming capacity are shown - the rest are at capacity"),
        layout_columns(
          col_widths = c(4, 2, 2, 4),
          div(
            div(class = "sb-section-label", "MW needed (minimum \u2013 ideal)"),
            range_slider("fc_need", 0, 100, c(10, 50), step = 1)
          ),
          selectInput("fc_when", "Needed by", choices = NULL),
          div(
            div(class = "sb-section-label", "Scope"),
            div(class = "seg-toggle",
                radioButtons("fc_scope", NULL,
                             choices = c("US" = "us", "Global" = "global"),
                             selected = "global", inline = TRUE))
          ),
          div(
            conditionalPanel(
              condition = "input.fc_scope == 'us'",
              selectizeInput("fc_state", "State", choices = NULL, multiple = TRUE,
                             options = list(placeholder = "All states",
                                            plugins = list("remove_button")))
            ),
            conditionalPanel(
              condition = "input.fc_scope == 'global'",
              selectizeInput("fc_country", "Country", choices = NULL, multiple = TRUE,
                             options = list(placeholder = "All countries",
                                            plugins = list("remove_button")))
            )
          )
        ),
        layout_columns(
          col_widths = c(3, 3, 3, 3),
          numericInput("fc_price", "Max $/kW (optional)", value = NA, min = 0),
          numericInput("fc_pue", "Max PUE (optional)", value = NA, min = 1, step = 0.05),
          selectInput("fc_sort", "Sort by",
                      c("Best match", "Most MW", "Lowest price")),
          div(class = "pt-4",
              checkboxInput("fc_near", "Include near misses", value = TRUE))
        )
      ),
      
      uiOutput("fc_summary"),
      uiOutput("fc_cards"),
      
      layout_columns(
        col_widths = c(7, 5),
        card(
          card_title("list", "All matches",
                     "Capacity arriving by your deadline, per site"),
          DTOutput("fc_table")
        ),
        card(
          full_screen = TRUE,
          card_title("map-location-dot", "Where they are"),
          leafletOutput("fc_map", height = 420)
        )
      )
    ),
    
    # --------------------------------------------------------
    # VERSION HISTORY
    # --------------------------------------------------------
    nav_panel(
      "Version History",
      
      card(
        card_title("rotate", "Refresh from DC.xlsx"),
        p(
          class = "version-note",
          "DC.xlsx is imported automatically every time the app starts. ",
          "Use this button only if you edit DC.xlsx while the app is running. ",
          "The current data is saved to version history first."
        ),
        div(actionButton("refresh_excel", "Refresh from DC.xlsx",
                         icon = icon("rotate"), class = "btn-primary")),
        br(),
        uiOutput("refresh_status")
      ),
      
      card(
        card_title("clipboard-check", "Data quality check",
                   "Capacity values in DC.xlsx that could not be read cleanly"),
        uiOutput("flag_summary"),
        DTOutput("flag_table"),
        card_footer(
          downloadButton(
            "download_flags",
            "Download review list (CSV)",
            class = "btn-outline-primary btn-sm"
          )
        )
      ),
      
      card(
        card_title("clock-rotate-left", "History"),
        p(
          class = "version-note",
          "Select a version to restore it. Restoring saves a new snapshot, ",
          "so you can always undo it."
        ),
        DTOutput("version_table")
      )
    )
  )
)

# ============================================================
# SERVER
# ============================================================

reset_view_button <- function(input_id) {
  easyButton(
    icon = "fa-globe",
    title = "Reset view",
    onClick = JS(sprintf(
      "function(btn, map) { Shiny.setInputValue('%s', Math.random(), {priority: 'event'}); }",
      input_id
    ))
  )
}

# Base map shared by all maps (data is added later through leafletProxy).
# Two basemaps with a Map / Satellite switch in the top-right corner:
#   "Map"       - CARTO Positron (the default every map opens on)
#   "Satellite" - Esri World Imagery + light place-name labels on top
# The switch is never touched by the marker/legend proxies, so the user's
# choice stays put while they filter and search.
base_leaflet <- function(...) {
  leaflet(options = leafletOptions(worldCopyJump = FALSE, minZoom = 2, maxZoom = 18, ...)) %>%
    addTiles(
      urlTemplate = CARTO_POSITRON_URL,
      attribution = CARTO_ATTRIBUTION,
      options = tileOptions(noWrap = TRUE),
      group = BASEMAP_DEFAULT
    ) %>%
    addProviderTiles(
      providers$Esri.WorldImagery,
      options = providerTileOptions(noWrap = TRUE, zIndex = 1),
      group = BASEMAP_SATELLITE
    ) %>%
    addTiles(
      urlTemplate = CARTO_LABELS_URL,
      attribution = CARTO_ATTRIBUTION,
      options = tileOptions(noWrap = TRUE, zIndex = 3, subdomains = "abcd"),
      group = BASEMAP_SATELLITE
    ) %>%
    addLayersControl(
      baseGroups = c(BASEMAP_DEFAULT, BASEMAP_SATELLITE),
      position = "topright",
      options = layersControlOptions(collapsed = FALSE)
    ) %>%
    # Guarantee every map opens on the regular map, never on satellite
    hideGroup(BASEMAP_SATELLITE) %>%
    showGroup(BASEMAP_DEFAULT) %>%
    setMaxBounds(lng1 = -180, lat1 = -85, lng2 = 180, lat2 = 85)
}

server <- function(input, output, session) {
  
  cur <- load_current()
  raw_data <- reactiveVal(cur)
  raw_pipeline <- reactiveVal(load_current_pipeline())
  
  cap_slider_max <- reactiveVal(100)
  pipe_slider_max <- reactiveVal(100)
  refresh_trigger <- reactiveVal(0)
  
  fmt <- function(x) format(round(x), big.mark = ",")
  
  # ----------------------------------------------------------
  # Current quarter (refreshed every few hours so a long-running
  # app rolls over on its own; downstream only re-runs on change)
  # ----------------------------------------------------------
  
  today_q <- reactiveVal(current_quarter_start())
  
  observe({
    invalidateLater(6 * 60 * 60 * 1000, session)
    q <- current_quarter_start()
    if (!identical(q, isolate(today_q()))) today_q(q)
  })
  
  # Upcoming capacity that has not slipped into the past
  pipeline_active <- reactive({
    raw_pipeline() %>% filter(!is.na(Quarter_Date), Quarter_Date >= today_q())
  })
  
  # What the Upcoming Capacity tab works from (past quarters are opt-in)
  pipeline_base <- reactive({
    if (isTRUE(input$pipeline_include_past)) raw_pipeline() else pipeline_active()
  })
  
  # Quarter labels available in the sheet, in date order
  quarters_of <- function(df) {
    if (nrow(df) == 0) return(character(0))
    df %>%
      filter(!is.na(Quarter_Date)) %>%
      distinct(Quarter, Quarter_Date) %>%
      arrange(Quarter_Date) %>%
      pull(Quarter)
  }
  
  quarter_choices <- reactive(quarters_of(pipeline_base()))
  
  # ----------------------------------------------------------
  # Header: "last updated" pill + data-quality pill
  # ----------------------------------------------------------
  
  last_update <- reactive({
    refresh_trigger()
    v <- list_versions() %>% filter(!grepl("^Auto-snapshot", note))
    if (nrow(v) == 0) NULL else v[1, ]
  })
  
  output$hdr_fresh <- renderUI({
    v <- last_update()
    
    if (is.null(v)) {
      return(span(class = "pill", "No data loaded yet"))
    }
    
    span(
      class = "pill", title = "When the data was last imported",
      span(class = "live-dot"),
      paste0("Updated ", substr(v$timestamp, 1, 16), " (v", v$version, ")")
    )
  })
  
  import_flags <- reactive({
    refresh_trigger()
    load_import_flags()
  })
  
  output$hdr_flags <- renderUI({
    n <- nrow(import_flags())
    if (n == 0) return(NULL)
    
    actionLink(
      "go_quality",
      label = tagList(
        icon("triangle-exclamation"),
        paste0(n, if (n == 1) " entry" else " entries", " to review")
      ),
      class = "pill pill-warn",
      title = "Some capacity values in DC.xlsx could not be read cleanly. Click to review."
    )
  })
  
  observeEvent(input$go_quality, {
    bslib::nav_select("main_tabs", "Version History", session = session)
  })
  
  output$data_freshness <- renderText({
    v <- last_update()
    if (is.null(v)) return("No data loaded yet")
    paste0("Data refreshed ", substr(v$timestamp, 1, 16), " \u00b7 v", v$version)
  })
  
  # ----------------------------------------------------------
  # Search
  # ----------------------------------------------------------
  
  search_raw_d <- debounce(
    reactive({
      s <- input$global_search
      if (is.null(s)) "" else trimws(s)
    }),
    250
  )
  
  apply_search <- function(df, cols) {
    term <- search_raw_d()
    
    if (!nzchar(term) || nrow(df) == 0) return(df)
    
    tokens <- parse_search_terms(term)
    if (length(tokens) == 0) return(df)
    
    cols <- intersect(cols, names(df))
    
    hay <- do.call(paste, c(
      lapply(df[cols], function(x) str_to_lower(coalesce(as.character(x), ""))),
      sep = " | "
    ))
    
    keep <- Reduce(`&`, lapply(tokens, function(t) str_detect(hay, fixed(t))))
    
    df[keep, ]
  }
  
  # Autocomplete suggestions
  observeEvent(list(raw_data(), raw_pipeline()), {
    cur_df <- raw_data()
    pl_df <- raw_pipeline()
    
    pick <- function(x, type) {
      x <- unique(as.character(x))
      x <- sort(x[!is.na(x) & nzchar(x)])
      tibble(label = x, type = rep(type, length(x)))
    }
    
    pl_city <- pl_df$City_clean[
      !is.na(pl_df$City_clean) &
        !str_ends(pl_df$City_clean, fixed(" (no city provided)"))
    ]
    
    sug <- bind_rows(
      pick(c(cur_df$Operator, pl_df$Operator), "Operator"),
      pick(c(cur_df$City, pl_city), "City"),
      pick(c(cur_df$State, pl_df$State), "State"),
      pick(c(cur_df$Country, pl_df$Country), "Country")
    )
    
    session$sendCustomMessage(
      "search_suggest",
      list(labels = unname(sug$label), types = unname(sug$type))
    )
  }, ignoreNULL = FALSE)
  
  observeEvent(input$search_selection, {
    s <- input$search_selection
    req(s$value, s$type)
    
    value <- s$value
    type  <- s$type
    
    # ---- Upcoming Capacity tab: use the pipeline's own filters ----
    if (identical(input$main_tabs, "Upcoming Capacity")) {
      pl <- raw_pipeline()
      us_states <- unique(pl$State[pl$is_us & !is.na(pl$State)])
      
      if (identical(type, "State") && value %in% us_states) {
        updateRadioButtons(session, "pipeline_scope", selected = "us")
        updateSelectizeInput(session, "pipeline_state_filter", selected = value)
        updateTextInput(session, "global_search", value = "")
        
      } else if (identical(type, "Country")) {
        updateRadioButtons(session, "pipeline_scope", selected = "global")
        updateSelectizeInput(session, "pipeline_country_filter", selected = value)
        updateTextInput(session, "global_search", value = "")
        
      } else {
        # Operator / City (and non-US states) have no dropdown on this tab,
        # so keep the selection as an exact-phrase text search instead.
        updateTextInput(session, "global_search", value = paste0('"', value, '"'))
      }
      return()
    }
    
    # ---- Data Centers tab ----
    if (identical(type, "State")) {
      updateSelectizeInput(session, "state_filter", selected = value)
    } else if (identical(type, "Country")) {
      updateSelectizeInput(session, "country_filter", selected = value)
    } else if (identical(type, "City")) {
      updateSelectizeInput(session, "city_filter", selected = value)
    } else if (identical(type, "Operator")) {
      updateSelectizeInput(session, "operator_filter", selected = value)
    }
    
    updateTextInput(session, "global_search", value = "")
  })
  
  # ----------------------------------------------------------
  # Location filter
  # ----------------------------------------------------------
  
  apply_location_filter <- function(df) {
    if (input$scope == "us") {
      df <- df %>% filter(is_us)
      
      if (length(input$state_filter) > 0) {
        df <- df %>% filter(State %in% input$state_filter)
      }
      
    } else if (length(input$state_filter) > 0 ||
               length(input$country_filter) > 0) {
      keep <- rep(FALSE, nrow(df))
      
      if (length(input$state_filter) > 0) keep <- keep | (df$State %in% input$state_filter)
      if (length(input$country_filter) > 0) keep <- keep | (df$Country %in% input$country_filter)
      
      df <- df[keep, ]
    }
    
    df
  }
  
  # ----------------------------------------------------------
  # Filter choices
  # ----------------------------------------------------------
  
  observeEvent(list(raw_data(), input$scope), {
    df <- raw_data()
    
    updateSelectizeInput(
      session, "state_filter",
      choices = sort(unique(df$State[df$is_us & !is.na(df$State)])),
      server = TRUE
    )
    
    updateSelectizeInput(
      session, "country_filter",
      choices = sort(unique(df$Country[!df$is_us & !is.na(df$Country)])),
      server = TRUE
    )
    
    df_scoped <- if (input$scope == "us") df %>% filter(is_us) else df
    
    max_cap <- suppressWarnings(max(df_scoped$Capacity_MW_est, na.rm = TRUE))
    if (!is.finite(max_cap)) max_cap <- 100
    
    cap_slider_max(ceiling(max_cap))
    
    updateSliderInput(
      session, "capacity_filter",
      max = ceiling(max_cap),
      value = c(0, ceiling(max_cap))
    )
  }, ignoreNULL = FALSE)
  
  observeEvent(raw_pipeline(), {
    countries <- raw_pipeline() %>%
      filter(!is.na(Country), Country != "") %>%
      distinct(Country) %>%
      arrange(Country) %>%
      pull(Country)
    
    updateSelectizeInput(session, "pipeline_country_filter", choices = countries, server = TRUE)
    
    pl_states <- raw_pipeline() %>%
      filter(is_us, !is.na(State), State != "") %>%
      distinct(State) %>%
      arrange(State) %>%
      pull(State)
    
    updateSelectizeInput(session, "pipeline_state_filter", choices = pl_states, server = TRUE)
  }, ignoreNULL = FALSE)
  
  # Quarter dropdown: built from the sheet, always in date order
  observeEvent(quarter_choices(), {
    qs <- quarter_choices()
    
    updateSelectizeInput(
      session, "quarter_filter",
      choices = qs,
      selected = intersect(input$quarter_filter, qs)
    )
  }, ignoreNULL = FALSE)
  
  # "Needed by" dropdown on Find Capacity (future quarters only).
  # Keeps the current choice if it still exists, otherwise ~1 year out.
  observeEvent(pipeline_active(), {
    qs <- quarters_of(pipeline_active())
    
    if (length(qs) == 0) {
      updateSelectInput(session, "fc_when", choices = character(0))
      return()
    }
    
    sel <- input$fc_when
    if (is.null(sel) || !sel %in% qs) {
      target <- seq(today_q(), by = "3 months", length.out = 5)[5]
      sel <- qs[which.min(abs(as.numeric(quarter_to_date(qs) - target)))]
    }
    
    updateSelectInput(session, "fc_when", choices = qs, selected = sel)
  }, ignoreNULL = FALSE)
  
  observeEvent(
    list(raw_data(), input$scope, input$state_filter, input$country_filter),
    {
      df <- apply_location_filter(raw_data())
      
      freezeReactiveValue(input, "city_filter")
      
      updateSelectizeInput(
        session, "city_filter",
        choices = sort(unique(df$City_clean)),
        server = TRUE
      )
    },
    ignoreNULL = FALSE
  )
  
  observeEvent(
    list(raw_data(), input$scope, input$state_filter,
         input$country_filter, input$city_filter),
    {
      df <- apply_location_filter(raw_data())
      
      if (length(input$city_filter) > 0) {
        df <- df %>% filter(City_clean %in% input$city_filter)
      }
      
      freezeReactiveValue(input, "operator_filter")
      
      updateSelectizeInput(
        session, "operator_filter",
        choices = sort(unique(df$Operator)),
        server = TRUE
      )
    },
    ignoreNULL = FALSE
  )
  
  # ----------------------------------------------------------
  # Sidebar helpers
  # ----------------------------------------------------------
  
  output$sb_summary <- renderUI({
    if (identical(input$main_tabs, "Find Capacity")) return(NULL)
    
    if (identical(input$main_tabs, "Upcoming Capacity")) {
      n <- nrow(filtered_pipeline())
      total <- nrow(pipeline_base())
      label <- "upcoming entries"
    } else {
      n <- nrow(filtered())
      total <- nrow(raw_data())
      label <- "sites"
    }
    
    div(
      class = "sb-summary",
      div(class = "sb-summary-main",
          tags$b(fmt(n)), paste0(" of ", fmt(total), " ", label))
    )
  })
  
  output$search_status <- renderUI({
    term <- search_raw_d()
    
    if (!nzchar(term)) return(NULL)
    
    on_pipeline <- identical(input$main_tabs, "Upcoming Capacity")
    
    n <- if (on_pipeline) nrow(filtered_pipeline()) else nrow(filtered())
    
    noun <- if (on_pipeline) c("upcoming entry", "upcoming entries") else c("site", "sites")
    
    shown_term <- gsub('"', "", term, fixed = TRUE)
    
    if (n == 0) {
      div(
        class = "sb-search-status is-empty",
        icon("circle-exclamation"),
        span("No ", noun[2], " match ", tags$b(shown_term),
             ". Try fewer words or reset the other filters.")
      )
    } else {
      div(
        class = "sb-search-status",
        span(class = "sb-dot"),
        span(tags$b(fmt(n)), " ", if (n == 1) noun[1] else noun[2],
             " match ", tags$b(shown_term))
      )
    }
  })
  
  output$filter_chips <- renderUI({
    make <- function(type, vals) {
      lapply(vals, function(v) {
        tags$span(
          class = "sb-chip", v,
          tags$a(
            class = "sb-chip-x", href = "#", HTML("&times;"),
            onclick = sprintf(
              "Shiny.setInputValue('remove_chip', {type:'%s', value:%s}, {priority:'event'}); return false;",
              type, jsonlite::toJSON(v, auto_unbox = TRUE)
            )
          )
        )
      })
    }
    
    chips <- c(
      make("state", input$state_filter),
      make("country", input$country_filter),
      make("city", input$city_filter),
      make("operator", input$operator_filter)
    )
    
    if (length(chips) == 0) return(NULL)
    
    div(class = "sb-chips", chips)
  })
  
  observeEvent(input$remove_chip, {
    id <- switch(
      input$remove_chip$type,
      state = "state_filter",
      country = "country_filter",
      city = "city_filter",
      operator = "operator_filter"
    )
    
    updateSelectizeInput(
      session, id,
      selected = setdiff(input[[id]], input$remove_chip$value)
    )
  })
  
  # Highlight toggle
  highlight_on <- reactiveVal(FALSE)
  
  observeEvent(input$highlight_toggle, {
    highlight_on(!highlight_on())
  })
  
  observeEvent(highlight_on(), {
    session$sendCustomMessage("hl_state", highlight_on())
  }, ignoreInit = TRUE)
  
  # Slider <-> typed number boxes
  sync_range <- function(id, max_fn) {
    lo_id <- paste0(id, "_lo")
    hi_id <- paste0(id, "_hi")
    
    observeEvent(input[[id]], {
      updateNumericInput(session, lo_id, value = input[[id]][1])
      updateNumericInput(session, hi_id, value = input[[id]][2])
    })
    
    observeEvent(list(input[[lo_id]], input[[hi_id]]), {
      lo <- input[[lo_id]]
      hi <- input[[hi_id]]
      req(is.numeric(lo), is.numeric(hi), is.finite(lo), is.finite(hi))
      
      mx <- max_fn()
      a <- min(max(min(lo, hi), 0), mx)
      b <- min(max(max(lo, hi), 0), mx)
      
      if (a != lo || b != hi) {
        updateNumericInput(session, lo_id, value = a)
        updateNumericInput(session, hi_id, value = b)
      }
      
      cur_val <- input[[id]]
      if (is.null(cur_val) || abs(cur_val[1] - a) > 1e-9 || abs(cur_val[2] - b) > 1e-9) {
        updateSliderInput(session, id, value = c(a, b))
      }
    }, ignoreInit = TRUE)
  }
  
  sync_range("capacity_filter", cap_slider_max)
  sync_range("pipeline_mw_filter", pipe_slider_max)
  
  # ----------------------------------------------------------
  # Reset
  # ----------------------------------------------------------
  
  observeEvent(input$reset_filters, {
    updateRadioButtons(session, "scope", selected = "global")
    updateRadioButtons(session, "pipeline_scope", selected = "global")
    updateRadioButtons(session, "donut_by", selected = "country")
    updateCheckboxInput(session, "pipeline_include_past", value = FALSE)
    
    for (id in c("state_filter", "country_filter", "city_filter",
                 "operator_filter", "pipeline_country_filter",
                 "pipeline_state_filter", "quarter_filter")) {
      updateSelectizeInput(session, id, selected = character(0))
    }
    
    highlight_on(FALSE)
    updateTextInput(session, "global_search", value = "")
    
    gmax <- suppressWarnings(max(raw_data()$Capacity_MW_est, na.rm = TRUE))
    if (!is.finite(gmax)) gmax <- 100
    
    updateSliderInput(session, "capacity_filter",
                      max = ceiling(gmax), value = c(0, ceiling(gmax)))
    
    pmax_mw <- suppressWarnings(max(raw_pipeline()$MW_available, na.rm = TRUE))
    if (!is.finite(pmax_mw)) pmax_mw <- 100
    
    updateSliderInput(session, "pipeline_mw_filter",
                      max = ceiling(pmax_mw), value = c(0, ceiling(pmax_mw)))
  })
  
  # ----------------------------------------------------------
  # Current-site filtering
  # ----------------------------------------------------------
  
  filtered <- reactive({
    df <- apply_location_filter(raw_data())
    
    if (length(input$city_filter) > 0) {
      df <- df %>% filter(City_clean %in% input$city_filter)
    }
    
    if (length(input$operator_filter) > 0) {
      df <- df %>% filter(Operator %in% input$operator_filter)
    }
    
    df <- apply_search(df, c("Operator", "City_clean", "State", "Country",
                             "Region", "Cooling", "Notes", "Contacts"))
    
    df %>%
      filter(
        is.na(Capacity_MW_est) |
          (Capacity_MW_est >= input$capacity_filter[1] &
             Capacity_MW_est <= input$capacity_filter[2])
      )
  })
  
  # ----------------------------------------------------------
  # Pipeline filtering (independent of the Data Centers filters;
  # the search bar applies to both tabs)
  # ----------------------------------------------------------
  
  filtered_pipeline <- reactive({
    df <- pipeline_base()
    
    if (identical(input$pipeline_scope, "us")) {
      df <- df %>% filter(is_us)
      
      if (length(input$pipeline_state_filter) > 0) {
        df <- df %>% filter(State %in% input$pipeline_state_filter)
      }
    } else if (length(input$pipeline_country_filter) > 0) {
      df <- df %>% filter(Country %in% input$pipeline_country_filter)
    }
    
    if (length(input$quarter_filter) > 0) {
      df <- df %>% filter(Quarter %in% input$quarter_filter)
    }
    
    if (length(input$pipeline_mw_filter) > 0) {
      df <- df %>%
        filter(
          MW_available >= input$pipeline_mw_filter[1] &
            MW_available <= input$pipeline_mw_filter[2]
        )
    }
    
    df <- apply_search(df, c("Operator", "City_clean", "State", "Country",
                             "Region", "Quarter"))
    
    df %>% arrange(Quarter_Date, Operator, Country, State, City_clean)
  })
  
  observeEvent(raw_pipeline(), {
    max_mw <- suppressWarnings(max(raw_pipeline()$MW_available, na.rm = TRUE))
    if (!is.finite(max_mw)) max_mw <- 100
    
    pipe_slider_max(ceiling(max_mw))
    
    updateSliderInput(session, "pipeline_mw_filter",
                      max = ceiling(max_mw), value = c(0, ceiling(max_mw)))
  }, ignoreNULL = FALSE)
  
  # ---- Year shortcut buttons: one per year found in the sheet, additive ----
  
  output$year_presets <- renderUI({
    qs <- quarter_choices()
    if (length(qs) == 0) return(NULL)
    
    yrs <- format(quarter_to_date(qs), "%Y")
    selected <- input$quarter_filter %||% character(0)
    
    div(
      class = "sb-presets",
      lapply(unique(yrs), function(y) {
        year_qs <- qs[yrs == y]
        is_active <- length(selected) > 0 && all(year_qs %in% selected)
        
        tags$button(
          type = "button",
          class = paste("btn btn-sm btn-outline-primary", if (is_active) "is-active"),
          `aria-pressed` = if (is_active) "true" else "false",
          onclick = sprintf(
            "Shiny.setInputValue('year_toggle', {year: '%s'}, {priority: 'event'});", y
          ),
          y
        )
      })
    )
  })
  
  observeEvent(input$year_toggle, {
    y <- input$year_toggle$year
    qs <- quarter_choices()
    year_qs <- qs[format(quarter_to_date(qs), "%Y") == y]
    selected <- input$quarter_filter %||% character(0)
    
    new_sel <- if (all(year_qs %in% selected)) {
      setdiff(selected, year_qs)
    } else {
      union(selected, year_qs)
    }
    
    updateSelectizeInput(session, "quarter_filter", selected = qs[qs %in% new_sel])
  })
  
  output$past_note <- renderUI({
    if (isTRUE(input$pipeline_include_past)) return(NULL)
    
    pl <- raw_pipeline()
    n_past <- sum(!is.na(pl$Quarter_Date) & pl$Quarter_Date < today_q())
    
    if (n_past == 0) return(NULL)
    
    div(
      class = "sb-note",
      icon("clock-rotate-left"),
      span(paste0(
        n_past, if (n_past == 1) " entry dated" else " entries dated",
        " before ", quarter_label(today_q()), " ",
        if (n_past == 1) "is" else "are", " hidden."
      ))
    )
  })
  
  # ----------------------------------------------------------
  # "Coming soon" highlight logic
  # ----------------------------------------------------------
  
  # The highlight window runs from the current quarter through the quarter
  # picked on the timeline strip (input$hl_through, e.g. "Q4 2027").
  hl_end <- reactive({
    q <- input$hl_through
    d <- if (is.null(q)) as.Date(NA) else quarter_to_date(q)
    if (is.na(d)) d <- seq(today_q(), by = "3 months", length.out = 8)[8]
    d
  })
  
  hl_label <- reactive({
    q <- input$hl_through
    if (is.null(q)) "within 8 quarters" else paste("through", q)
  })
  
  upcoming_soon_keys <- reactive({
    pl <- pipeline_active()
    
    if (nrow(pl) == 0) return(character(0))
    
    end_q <- hl_end()
    
    pl %>%
      filter(Quarter_Date <= end_q) %>%
      mutate(key = paste(Operator, City_clean, State, Country, sep = "|")) %>%
      pull(key) %>%
      unique()
  })
  
  # Timeline strip: one bar per quarter from now to the last quarter in the data.
  # Built from the data, so it never needs manual updating. Clicking/dragging is
  # handled in JavaScript (no re-render), which reports the choice as input$hl_through.
  output$hl_timeline <- renderUI({
    pl <- pipeline_active()
    if (nrow(pl) == 0) return(NULL)
    
    last_q <- max(pl$Quarter_Date, na.rm = TRUE)
    qs <- seq(today_q(), last_q, by = "3 months")
    
    q_df <- tibble(Quarter_Date = qs, Quarter = quarter_label(qs)) %>%
      left_join(
        pl %>%
          group_by(Quarter_Date) %>%
          summarise(MW = sum(MW_available, na.rm = TRUE), .groups = "drop"),
        by = "Quarter_Date"
      ) %>%
      mutate(MW = coalesce(MW, 0))
    
    n <- nrow(q_df)
    mx <- max(q_df$MW, 1)
    default_idx <- min(7L, n - 1L)
    
    cols <- lapply(seq_len(n), function(i) {
      r <- q_df[i, ]
      qn <- substr(r$Quarter, 1, 2)
      yr <- substr(r$Quarter, 6, 7)
      show_year <- i == 1 || qn == "Q1"
      
      tags$button(
        type = "button", class = "tl-col",
        `data-q` = r$Quarter, `data-mw` = round(r$MW, 1),
        title = paste0(r$Quarter, " \u00b7 ", fmt(r$MW), " MW arriving"),
        `aria-label` = paste0("Highlight through ", r$Quarter, ", ", fmt(r$MW), " MW arriving"),
        div(class = "tl-bar",
            div(class = "tl-fill",
                style = sprintf("height:%d%%", as.integer(max(6, round(100 * r$MW / mx)))))),
        span(class = "tl-q", qn),
        span(class = "tl-y", if (show_year) paste0("'", yr) else HTML("&nbsp;"))
      )
    })
    
    div(
      id = "hl_tl_wrap", class = "hl-tl", `data-default` = default_idx,
      div(
        class = "tl-head",
        span(class = "tl-title", "Highlight window"),
        div(class = "tl-chips",
            tags$button(type = "button", class = "tl-chip", `data-n` = "4", "4Q"),
            tags$button(type = "button", class = "tl-chip", `data-n` = "8", "8Q"),
            tags$button(type = "button", class = "tl-chip", `data-n` = "99", "All"))
      ),
      div(class = "tl-strip", cols),
      div(id = "hl_tl_read", class = "tl-read", "")
    )
  })
  
  # Count of sites in view with capacity arriving inside the chosen window
  output$hl_count <- renderText({
    keys <- upcoming_soon_keys()
    n <- filtered() %>%
      mutate(key = paste(Operator, City_clean, State, Country, sep = "|")) %>%
      filter(key %in% keys) %>%
      nrow()
    paste0(fmt(n), if (n == 1) " site" else " sites", " coming soon")
  })
  
  # ==========================================================
  # KPI tiles
  # ==========================================================
  
  output$vb_operators <- renderText(n_distinct(filtered()$Operator))
  output$vb_cities <- renderText(n_distinct(filtered()$City_clean))
  output$vb_capacity <- renderText(fmt(sum(filtered()$Capacity_MW_est, na.rm = TRUE)))
  
  output$vb_pipeline_entries <- renderText(nrow(filtered_pipeline()))
  output$vb_pipeline_sites <- renderText({
    filtered_pipeline() %>%
      distinct(Operator, City_clean, State, Country) %>%
      nrow()
  })
  output$vb_pipeline_mw <- renderText(fmt(sum(filtered_pipeline()$MW_available, na.rm = TRUE)))
  
  # ==========================================================
  # MAP READINESS
  #   Maps on tabs that have not been opened yet do not exist in the
  #   browser. Leaflet reports its zoom/bounds as soon as a map is
  #   built, so we only start sending proxy updates after that.
  # ==========================================================
  
  pipeline_map_ready <- reactiveVal(FALSE)
  fc_map_ready <- reactiveVal(FALSE)
  
  observe({
    if (!is.null(input$pipeline_map_zoom) || !is.null(input$pipeline_map_bounds)) {
      pipeline_map_ready(TRUE)
    }
  })
  
  observe({
    if (!is.null(input$fc_map_zoom) || !is.null(input$fc_map_bounds)) {
      fc_map_ready(TRUE)
    }
  })
  
  # ==========================================================
  # DATA CENTERS MAP
  # ==========================================================
  
  output$map <- renderLeaflet({
    base_leaflet() %>%
      setView(lng = -98.5, lat = 39.5, zoom = 4) %>%
      addEasyButton(reset_view_button("map_reset"))
  })
  
  reset_map_view <- function() {
    proxy <- leafletProxy("map")
    
    if (identical(input$scope, "us")) {
      proxy %>% setView(lng = -98.5, lat = 39.5, zoom = 4)
    } else {
      proxy %>% setView(lng = -20, lat = 25, zoom = 2)
    }
  }
  
  reset_pipeline_view <- function() {
    req(isTRUE(pipeline_map_ready()))
    proxy <- leafletProxy("pipeline_map")
    
    if (identical(input$pipeline_scope, "us")) {
      proxy %>% setView(lng = -98.5, lat = 39.5, zoom = 4)
    } else {
      proxy %>% setView(lng = -20, lat = 25, zoom = 2)
    }
  }
  
  observeEvent(input$pipeline_scope, reset_pipeline_view(), ignoreInit = TRUE)
  observeEvent(input$scope, reset_map_view(), ignoreInit = TRUE)
  observeEvent(input$map_reset, reset_map_view())
  observeEvent(input$pipeline_map_reset, reset_pipeline_view())
  
  # Searching zooms the map to the matching sites; clearing resets the view
  observeEvent(search_raw_d(), {
    req(identical(input$main_tabs, "Data Centers"))
    
    if (!nzchar(search_raw_d())) {
      reset_map_view()
      return()
    }
    
    pts <- filtered() %>% filter(!is.na(Latitude), !is.na(Longitude))
    
    if (nrow(pts) == 0) return()
    
    leafletProxy("map") %>% fit_points(pts$Latitude, pts$Longitude)
  }, ignoreInit = TRUE)
  
  # Markers: heat colour + size by MW, gray = no capacity info,
  # orange = arriving in the highlight window (when highlight is on)
  observe({
    df <- filtered() %>% filter(!is.na(Latitude), !is.na(Longitude))
    
    keys <- upcoming_soon_keys()
    
    # Spread rows that share identical coordinates into a small ring (~250 m)
    df <- df %>%
      arrange(Operator, City_clean) %>%
      group_by(Latitude, Longitude) %>%
      mutate(
        point_n = n(),
        point_i = row_number(),
        angle = if_else(point_n > 1, 2 * pi * (point_i - 1) / point_n, 0),
        off_km = if_else(point_n > 1, 0.25, 0),
        map_lat = Latitude + (off_km / 111.32) * sin(angle),
        map_lng = Longitude + (off_km / (111.32 * pmax(cos(Latitude * pi / 180), 0.2))) * cos(angle)
      ) %>%
      ungroup() %>%
      mutate(
        key = paste(Operator, City_clean, State, Country, sep = "|"),
        is_upcoming_soon = if (isTRUE(highlight_on())) key %in% keys else FALSE,
        
        has_capacity = !is.na(Capacity_MW_est) & Capacity_MW_est > 0,
        is_zero      = !is.na(Capacity_MW_est) & Capacity_MW_est == 0,
        no_info      = is.na(Capacity_MW_est),
        
        # Show the MW estimate line when it adds something to the raw text
        has_est = !is.na(Capacity_MW_est),
        show_est = has_est & (
          abs(coalesce(Capacity_MW_max, Capacity_MW_est) -
                coalesce(Capacity_MW_min, Capacity_MW_est)) >= 0.05 |
            !is.na(Capacity_Flag)
        ),
        est_txt = esc_html(mw_range_txt(Capacity_MW_est, Capacity_MW_min, Capacity_MW_max)),
        
        op_txt = esc_html(coalesce(Operator, "Unknown operator")),
        city_txt = esc_html(coalesce(City_clean, "")),
        place_txt = esc_html(
          if_else(is.na(State) | State == "", coalesce(Country, ""), State)
        ),
        cap_txt = esc_html(coalesce(Capacity, "No info")),
        cool_txt = esc_html(coalesce(Cooling, "Unknown")),
        notes_txt = esc_html(coalesce(Notes, "")),
        popup_html = paste0(
          "<div style='min-width:200px; max-width:280px'>",
          "<b>", op_txt, "</b><br>",
          city_txt, ", ", place_txt, "<br>",
          "Capacity: ", cap_txt,
          if_else(show_est, paste0("<br>Estimate: ", est_txt), ""),
          if_else(
            !is.na(Capacity_Flag),
            "<br><span style='color:#B45309'>&#9888; Capacity text needs review</span>",
            ""
          ),
          if_else(is.na(Cooling), "", paste0("<br>Cooling: ", cool_txt)),
          # Native <details>: the browser keeps it closed until it is clicked
          if_else(
            is.na(Notes) | Notes == "", "",
            paste0(
              "<details class='map-note'>",
              "<summary><span class='map-note-i'>i</span>Notes</summary>",
              "<div class='map-note-body'>", notes_txt, "</div>",
              "</details>"
            )
          ),
          if_else(
            is_upcoming_soon,
            "<br><b style='color:#F59E0B'>Capacity coming available soon - see Upcoming Capacity tab</b>",
            ""
          ),
          "</div>"
        ),
        label_txt = paste0(
          op_txt, " \u2014 ", city_txt,
          if_else(has_est, paste0(" \u00b7 ", est_txt), "")
        )
      )
    
    proxy <- leafletProxy("map") %>%
      clearMarkers() %>%
      clearMarkerClusters() %>%
      clearControls()
    
    if (nrow(df) == 0) return()
    
    cap_vals <- df$Capacity_MW_est[df$has_capacity]
    min_mw <- suppressWarnings(min(cap_vals, na.rm = TRUE))
    max_mw <- suppressWarnings(max(cap_vals, na.rm = TRUE))
    if (!is.finite(min_mw)) min_mw <- 0
    if (!is.finite(max_mw)) max_mw <- min_mw + 1
    if (min_mw == max_mw) max_mw <- min_mw + 1
    
    pal <- colorNumeric(
      palette = c("#DBEAFE", "#60A5FA", "#1D4ED8", "#0A1A4A"),
      domain = c(min_mw, max_mw),
      na.color = "#9CA3AF"
    )
    
    heat_val <- ifelse(df$has_capacity, df$Capacity_MW_est, NA_real_)
    
    df <- df %>%
      mutate(
        radius = if_else(
            has_capacity,
            6 + 12 * sqrt(pmax(0, (coalesce(Capacity_MW_est, min_mw) - min_mw) / (max_mw - min_mw))),
            6
          ),
        fill_col = case_when(
          is_upcoming_soon ~ "#F59E0B",
          has_capacity     ~ pal(heat_val),
          is_zero          ~ "#E5E7EB",
          TRUE             ~ "#9CA3AF"
        ),
        border_col = case_when(
          is_upcoming_soon ~ "#FCD34D",
          has_capacity     ~ "#BFDBFE",
          TRUE             ~ "#D1D5DB"
        )
      ) %>%
      arrange(desc(has_capacity), desc(Capacity_MW_est))
    
    proxy %>%
      addCircleMarkers(
        data = df,
        lng = ~map_lng,
        lat = ~map_lat,
        radius = ~radius,
        stroke = TRUE,
        weight = 1.5,
        color = ~border_col,
        fillColor = ~fill_col,
        fillOpacity = 0.85,
        popup = ~popup_html,
        label = ~lapply(label_txt, htmltools::HTML),
        clusterOptions = markerClusterOptions(
          disableClusteringAtZoom = 14,
          spiderfyOnMaxZoom = FALSE,
          maxClusterRadius = 40,
          iconCreateFunction = JS(CLUSTER_ICON_JS)
        )
      )
    
    any_gray <- any(df$no_info)
    any_orange <- any(df$is_upcoming_soon)
    
    legend_html <- paste0(
      "<div class='dc-legend'>",
      
      if (length(cap_vals) > 0) paste0(
        "<div class='dc-legend-title'>Current capacity</div>",
        "<div class='dc-legend-bar'></div>",
        "<div class='dc-legend-scale'><span>",
        format(round(min_mw), big.mark = ","), " MW</span><span>",
        format(round(max_mw), big.mark = ","), " MW</span></div>"
      ) else "",
      
      if (any_gray) paste0(
        "<div class='dc-legend-row'>",
        "<span class='dc-legend-dot' style='background:#9CA3AF'></span>",
        "No info</div>"
      ) else "",
      
      if (any_orange) paste0(
        "<div class='dc-legend-row'>",
        "<span class='dc-legend-dot' style='background:#F59E0B'></span>",
        "Coming online ", hl_label(), "</div>"
      ) else "",
      
      "</div>"
    )
    
    proxy %>% addControl(html = legend_html, position = "bottomright", className = "")
  })
  
  # ==========================================================
  # DATA CENTERS CHARTS
  # ==========================================================
  
  output$top_operators_chart <- renderPlotly({
    df <- filtered() %>%
      filter(!is.na(Capacity_MW_est), Capacity_MW_est > 0) %>%
      mutate(
        Operator = coalesce(Operator, "Unknown operator"),
        cap_lo = coalesce(Capacity_MW_min, Capacity_MW_est),
        cap_hi = coalesce(Capacity_MW_max, Capacity_MW_est)
      ) %>%
      group_by(Operator) %>%
      summarise(
        MW = sum(Capacity_MW_est, na.rm = TRUE),
        MW_lo = sum(cap_lo, na.rm = TRUE),
        MW_hi = sum(cap_hi, na.rm = TRUE),
        Sites = n(), .groups = "drop"
      ) %>%
      arrange(desc(MW)) %>%
      slice_head(n = 10)
    
    if (nrow(df) == 0) return(empty_plot("No capacity matches the current search and filters."))
    
    df <- df %>%
      mutate(
        Operator = factor(Operator, levels = rev(Operator)),
        span = mw_span_txt(MW_lo, MW_hi),
        hover = paste0(
          "<b>", Operator, "</b><br>", fmt(MW), " MW est. across ",
          Sites, if_else(Sites == 1L, " site", " sites"),
          if_else(nzchar(span), paste0("<br>Range: ", span, " MW"), "")
        )
      )
    
    plot_ly(
      df, x = ~MW, y = ~Operator, type = "bar", orientation = "h",
      text = ~fmt(MW), textposition = "outside", cliponaxis = FALSE,
      textfont = list(color = "#CBD5E1"),
      hovertext = ~hover,
      hoverinfo = "text",
      marker = list(color = "#3B82F6")
    ) %>%
      dark_plot(
        xaxis = list(title = NULL, gridcolor = "#1E2A3A", zeroline = FALSE,
                     tickfont = list(color = "#94A3B8")),
        yaxis = list(title = NULL, tickfont = list(color = "#CBD5E1")),
        margin = list(l = 10, r = 50, t = 10, b = 30),
        showlegend = FALSE
      )
  })
  
  # ---- Capacity share donut: by country OR by city ----
  
  # TRUE when everything in view belongs to a single country
  single_country <- reactive({
    n_distinct(na.omit(filtered()$Country)) == 1
  })
  
  # Auto-flip: one country in view -> city view; otherwise -> country view.
  # Fires only when that state changes, so a manual toggle isn't overridden.
  observeEvent(single_country(), {
    updateRadioButtons(session, "donut_by",
                       selected = if (single_country()) "city" else "country")
  }, ignoreInit = FALSE)
  
  output$donut_title <- renderText({
    if (identical(input$donut_by, "city")) "Capacity by city" else "Capacity by country"
  })
  
  output$share_chart <- renderPlotly({
    by_city <- identical(input$donut_by, "city")
    
    df <- filtered() %>%
      filter(!is.na(Capacity_MW_est), Capacity_MW_est > 0) %>%
      mutate(
        cap_lo = coalesce(Capacity_MW_min, Capacity_MW_est),
        cap_hi = coalesce(Capacity_MW_max, Capacity_MW_est)
      )
    
    df <- if (by_city) {
      df %>% mutate(Label = case_when(
        is.na(City_clean) ~ "Unknown",
        is_us & !is.na(State) & State != "" ~ paste0(City_clean, ", ", State),
        TRUE ~ City_clean
      ))
    } else {
      df %>% mutate(Label = coalesce(Country, "Unknown"))
    }
    
    df <- df %>%
      group_by(Label) %>%
      summarise(
        MW = sum(Capacity_MW_est, na.rm = TRUE),
        MW_lo = sum(cap_lo, na.rm = TRUE),
        MW_hi = sum(cap_hi, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(desc(MW))
    
    if (nrow(df) == 0) return(empty_plot("No capacity matches the current search and filters."))
    
    if (nrow(df) > 7) {
      tail_rows <- df[8:nrow(df), ]
      df <- bind_rows(
        df %>% slice_head(n = 7),
        tibble(Label = if (by_city) "Other cities" else "Other",
               MW = sum(tail_rows$MW),
               MW_lo = sum(tail_rows$MW_lo),
               MW_hi = sum(tail_rows$MW_hi))
      )
    }
    
    span <- mw_span_txt(df$MW_lo, df$MW_hi)
    df$range_line <- ifelse(nzchar(span), paste0("Range: ", span, " MW"), "")
    
    cols <- c("#1D4ED8", "#3B82F6", "#60A5FA", "#93C5FD", "#2DD4BF",
              "#F59E0B", "#A78BFA", "#64748B")[seq_len(nrow(df))]
    
    plot_ly(
      df, labels = ~Label, values = ~MW, type = "pie", hole = 0.62,
      sort = FALSE, direction = "clockwise",
      textinfo = "none",
      text = ~range_line,
      hovertemplate = "<b>%{label}</b><br>%{value:,.0f} MW (%{percent})<br>%{text}<extra></extra>",
      marker = list(colors = cols, line = list(color = "#0F1620", width = 2))
    ) %>%
      dark_plot(
        showlegend = TRUE,
        legend = list(font = list(color = "#CBD5E1", size = 12)),
        margin = list(l = 10, r = 10, t = 10, b = 10)
      )
  })
  
  # ==========================================================
  # UPCOMING CAPACITY MAP
  #   The base map is drawn once. Filters, search and tab changes
  #   only swap the markers through leafletProxy, so there is no
  #   flicker and the user's zoom/pan is never reset.
  # ==========================================================
  
  output$pipeline_map <- renderLeaflet({
    us <- identical(isolate(input$pipeline_scope), "us")
    
    base_leaflet() %>%
      setView(
        lng = if (us) -98.5 else -20,
        lat = if (us) 39.5 else 25,
        zoom = if (us) 4 else 2
      ) %>%
      addEasyButton(reset_view_button("pipeline_map_reset"))
  })
  
  # Points for the map: country centroid fallback + small ring so stacked
  # rows at the same spot can each be clicked
  pipeline_map_df <- reactive({
    df <- filtered_pipeline()
    
    if (nrow(df) == 0) return(df)
    
    df %>%
      left_join(
        country_centroids %>%
          rename(Country_Latitude = Latitude, Country_Longitude = Longitude),
        by = "Country"
      ) %>%
      mutate(
        map_base_lat = coalesce(Latitude, Country_Latitude),
        map_base_lng = coalesce(Longitude, Country_Longitude)
      ) %>%
      group_by(map_base_lat, map_base_lng) %>%
      mutate(
        point_n = n(),
        point_i = row_number(),
        angle = if_else(point_n > 1, 2 * pi * (point_i - 1) / point_n, 0),
        offset_km = if_else(point_n > 1, 8, 0),
        map_lat = map_base_lat + (offset_km / 111.32) * sin(angle),
        map_lng = map_base_lng +
          (offset_km / (111.32 * pmax(cos(map_base_lat * pi / 180), 0.2))) * cos(angle)
      ) %>%
      ungroup() %>%
      filter(!is.na(map_lat), !is.na(map_lng), !is.na(MW_available), MW_available >= 0)
  })
  
  observe({
    req(isTRUE(pipeline_map_ready()))
    
    map_df <- pipeline_map_df()
    
    proxy <- leafletProxy("pipeline_map") %>%
      clearMarkers() %>%
      clearControls()
    
    if (nrow(map_df) == 0) return()
    
    min_mw <- min(map_df$MW_available, na.rm = TRUE)
    max_mw <- max(map_df$MW_available, na.rm = TRUE)
    
    if (!is.finite(min_mw)) min_mw <- 0
    if (!is.finite(max_mw)) max_mw <- 1
    
    pal <- colorNumeric(
      palette = c("#DBEAFE", "#60A5FA", "#1D4ED8", "#0A1A4A"),
      domain = if (min_mw == max_mw) c(min_mw, min_mw + 1) else c(min_mw, max_mw),
      na.color = "#64748B"
    )
    
    radius <- if (max_mw == min_mw) {
      rep(9, nrow(map_df))
    } else {
      6 + 12 * sqrt((map_df$MW_available - min_mw) / (max_mw - min_mw))
    }
    
    op_txt <- ifelse(is.na(map_df$Operator), "Unknown operator", esc_html(map_df$Operator))
    place_txt <- ifelse(
      is.na(map_df$City_clean) | map_df$City_clean == "",
      esc_html(map_df$Country),
      esc_html(map_df$City_clean)
    )
    state_txt <- ifelse(
      !is.na(map_df$State) & map_df$State != "",
      paste0(", ", esc_html(map_df$State)),
      ""
    )
    mw_txt <- esc_html(mw_range_txt(map_df$MW_available, map_df$MW_min, map_df$MW_max))
    
    popup_text <- paste0(
      "<div style='min-width:210px'>",
      "<b style='font-size:15px'>", op_txt, "</b><br>",
      place_txt, state_txt,
      "<br><br>",
      "<b>Coming online:</b> ", esc_html(map_df$Quarter), "<br>",
      "<b>MW available:</b> ", mw_txt,
      "</div>"
    )
    
    label_text <- paste0(
      op_txt, " \u2014 ", place_txt, " \u00b7 ", esc_html(map_df$Quarter),
      " \u00b7 ", mw_txt
    )
    
    proxy %>%
      addCircleMarkers(
        lng = map_df$map_lng,
        lat = map_df$map_lat,
        radius = radius,
        stroke = TRUE,
        weight = 1.5,
        color = "#BFDBFE",
        fillColor = pal(map_df$MW_available),
        fillOpacity = 0.85,
        popup = popup_text,
        label = lapply(label_text, htmltools::HTML)
      ) %>%
      addLegend(
        position = "bottomright",
        pal = pal,
        values = map_df$MW_available,
        title = "Upcoming MW (est.)",
        opacity = 0.9,
        labFormat = labelFormat(suffix = " MW")
      )
  })
  
  # Searching zooms to the matching entries; clearing the search resets the view
  observeEvent(search_raw_d(), {
    req(identical(input$main_tabs, "Upcoming Capacity"), isTRUE(pipeline_map_ready()))
    
    if (!nzchar(search_raw_d())) {
      reset_pipeline_view()
      return()
    }
    
    pts <- pipeline_map_df()
    if (nrow(pts) == 0) return()
    
    leafletProxy("pipeline_map") %>% fit_points(pts$map_lat, pts$map_lng)
  }, ignoreInit = TRUE)
  
  # Picking states / countries zooms to them; clearing them resets the view
  observeEvent(
    list(input$pipeline_state_filter, input$pipeline_country_filter),
    {
      req(identical(input$main_tabs, "Upcoming Capacity"), isTRUE(pipeline_map_ready()))
      
      if (length(c(input$pipeline_state_filter, input$pipeline_country_filter)) == 0) {
        reset_pipeline_view()
        return()
      }
      
      pts <- pipeline_map_df()
      if (nrow(pts) == 0) return()
      
      leafletProxy("pipeline_map") %>% fit_points(pts$map_lat, pts$map_lng)
    },
    ignoreInit = TRUE, ignoreNULL = FALSE
  )
  
  # ==========================================================
  # CURRENT TABLE (with notes buttons)
  # ==========================================================
  
  table_data <- reactive({
    filtered() %>%
      select(
        Operator,
        City = City_clean,
        State,
        Country,
        Region,
        Capacity,
        `Capacity (MW est.)` = Capacity_MW_est,
        `MW Min` = Capacity_MW_min,
        `MW Max` = Capacity_MW_max,
        Notes
      ) %>%
      arrange(Operator, Country, State, City)
  })
  
  output$table <- renderDT({
    df <- table_data()
    
    has_note <- !is.na(df$Notes) & nzchar(df$Notes)
    
    # The (i) button carries the note text; the server turns a click into a modal.
    # Sub-title per row: "City, State"
    sub_txt <- mapply(function(ct, st) {
      paste(c(ct, st)[!is.na(c(ct, st)) & nzchar(c(ct, st))], collapse = ", ")
    }, df$City, df$State, USE.NAMES = FALSE)
    
    note_btn <- ifelse(
      has_note,
      sprintf(
        "<button type='button' class='note-btn' aria-label='View notes' title='View notes' data-title='%s' data-sub='%s' data-note='%s'>i</button>",
        htmltools::htmlEscape(coalesce(df$Operator, "Unknown operator"), attribute = TRUE),
        htmltools::htmlEscape(sub_txt, attribute = TRUE),
        htmltools::htmlEscape(coalesce(df$Notes, ""), attribute = TRUE)
      ),
      "<span class='note-none'>&ndash;</span>"
    )
    
    # Escape the text columns since the table renders HTML for the button column
    df <- df %>%
      mutate(across(
        where(is.character),
        ~ ifelse(is.na(.x), "", htmltools::htmlEscape(.x))
      ))
    
    df$Notes <- note_btn
    
    dt <- datatable(
      df,
      rownames = FALSE,
      escape = FALSE,
      class = "compact hover",
      selection = "none",
      options = list(
        pageLength = 15,
        dom = "lfrtip",
        columnDefs = list(list(
          targets = ncol(df) - 1,
          orderable = FALSE, searchable = FALSE,
          className = "dt-center", width = "60px"
        )),
        language = list(
          emptyTable = "No sites match the current search and filters.",
          zeroRecords = "No sites match the current search and filters."
        )
      )
    )
    
    add_color_bar(dt, df, "Capacity (MW est.)")
  })
  
  # Notes modal: only opens when an (i) button is clicked
  observeEvent(input$note_click, {
    n <- input$note_click
    req(n$note, nzchar(n$note))
    
    showModal(modalDialog(
      title = div(
        div(class = "note-modal-op", n$title),
        if (nzchar(n$sub)) div(class = "note-modal-sub", n$sub)
      ),
      div(class = "note-modal-body", n$note),
      easyClose = TRUE,
      size = "m",
      footer = modalButton("Close")
    ))
  })
  
  output$download_filtered <- downloadHandler(
    filename = function() {
      paste0("data_centers_filtered_", format(Sys.Date(), "%Y%m%d"), ".csv")
    },
    content = function(file) {
      write.csv(table_data() %>% select(-Notes), file, row.names = FALSE, na = "")
    }
  )
  
  # ==========================================================
  # UPCOMING CAPACITY BY QUARTER
  # ==========================================================
  
  output$pipeline_quarter_chart <- renderPlotly({
    df <- filtered_pipeline()
    
    if (nrow(df) == 0) {
      return(empty_plot("No upcoming capacity matches the current search and filters."))
    }
    
    df <- df %>%
      mutate(
        Operator = if_else(is.na(Operator) | Operator == "", "Unknown operator", Operator),
        MW_lo = coalesce(MW_min, MW_available),
        MW_hi = coalesce(MW_max, MW_available)
      )
    
    # Every quarter between the first and last one in view, so gaps show as empty
    all_dates <- seq(min(df$Quarter_Date), max(df$Quarter_Date), by = "3 months")
    
    quarter_totals <- df %>%
      group_by(Quarter) %>%
      summarise(
        MW = sum(MW_available, na.rm = TRUE),
        MW_lo = sum(MW_lo, na.rm = TRUE),
        MW_hi = sum(MW_hi, na.rm = TRUE),
        .groups = "drop"
      )
    
    operator_breakdown <- df %>%
      group_by(Quarter, Quarter_Date, Operator) %>%
      summarise(
        MW = sum(MW_available, na.rm = TRUE),
        MW_lo = sum(MW_lo, na.rm = TRUE),
        MW_hi = sum(MW_hi, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(
        span = mw_span_txt(MW_lo, MW_hi),
        line = paste0(
          "<b>", Operator, "</b>: ", fmt_mw(MW), " MW",
          if_else(nzchar(span), paste0(" (", span, ")"), "")
        )
      ) %>%
      arrange(Quarter_Date, desc(MW)) %>%
      group_by(Quarter) %>%
      summarise(Breakdown = paste0(line, collapse = "<br>"), .groups = "drop")
    
    chart_data <- tibble(Quarter = quarter_label(all_dates)) %>%
      left_join(quarter_totals, by = "Quarter") %>%
      left_join(operator_breakdown, by = "Quarter") %>%
      mutate(
        MW = coalesce(MW, 0),
        total_span = mw_span_txt(MW_lo, MW_hi),
        bar_label = if_else(MW > 0, format(round(MW, 1), big.mark = ","), ""),
        hover = paste0(
          "<b>", Quarter, "</b>",
          "<br><b>Total: ", format(round(MW, 1), big.mark = ","), " MW</b>",
          if_else(!is.na(total_span) & nzchar(total_span),
                  paste0("<br>Range: ", total_span, " MW"), ""),
          if_else(is.na(Breakdown), "", paste0("<br><br>", Breakdown))
        )
      )
    
    plot_ly(
      data = chart_data,
      x = ~Quarter, y = ~MW, type = "bar",
      text = ~bar_label,
      textposition = "outside", cliponaxis = FALSE,
      textfont = list(color = "#CBD5E1"),
      hovertext = ~hover,
      hoverinfo = "text",
      marker = list(color = "#3B82F6")
    ) %>%
      dark_plot(
        xaxis = list(
          title = NULL, type = "category",
          categoryorder = "array", categoryarray = chart_data$Quarter,
          tickfont = list(color = "#CBD5E1"), gridcolor = "#1E2A3A", linecolor = "#2A3A50"
        ),
        yaxis = list(
          title = list(text = "Upcoming capacity (MW)", font = list(color = "#CBD5E1")),
          tickfont = list(color = "#CBD5E1"), gridcolor = "#1E2A3A", zerolinecolor = "#2A3A50"
        ),
        hovermode = "closest",
        margin = list(l = 65, r = 30, t = 30, b = 60),
        showlegend = FALSE
      )
  })
  
  # ==========================================================
  # UPCOMING CAPACITY TABLE
  # ==========================================================
  
  pipeline_table_data <- reactive({
    filtered_pipeline() %>%
      arrange(Quarter_Date, Operator, Country, State, City_clean) %>%
      transmute(
        Operator,
        City = City_clean,
        State,
        Country,
        Region,
        Quarter,
        `MW Available` = MW_available,
        `MW Min` = MW_min,
        `MW Max` = MW_max,
        Quarter_Order = as.integer(Quarter_Date)
      )
  })
  
  output$pipeline_table <- renderDT({
    df <- pipeline_table_data()
    
    if (nrow(df) == 0) {
      return(datatable(
        tibble(Message = "No upcoming capacity matches the current search and filters."),
        rownames = FALSE,
        options = list(dom = "t")
      ))
    }
    
    order_col <- which(names(df) == "Quarter_Order") - 1L
    quarter_col <- which(names(df) == "Quarter") - 1L
    
    dt <- datatable(
      df,
      rownames = FALSE,
      class = "compact hover",
      width = "100%",
      selection = "none",
      options = list(
        pageLength = 25,
        dom = "lfrtip",
        autoWidth = FALSE,
        order = list(list(order_col, "asc")),
        columnDefs = list(
          list(visible = FALSE, targets = order_col),
          list(orderData = order_col, targets = quarter_col))
      )
    )
    
    add_color_bar(dt, df, "MW Available")
  })
  
  output$download_pipeline <- downloadHandler(
    filename = function() {
      paste0("upcoming_capacity_filtered_", format(Sys.Date(), "%Y%m%d"), ".csv")
    },
    content = function(file) {
      pipeline_table_data() %>%
        select(-Quarter_Order) %>%
        write.csv(file, row.names = FALSE, na = "")
    }
  )

  # Client PDF: same filtered rows as the table, operator names left out
  output$download_pipeline_pdf <- downloadHandler(
    filename = function() {
      paste0("ByteBridge_Upcoming_Capacity_", format(Sys.Date(), "%Y%m%d"), ".pdf")
    },
    content = function(file) {
      full <- pipeline_table_data()

      # Operator names: dropped as a column AND scrubbed from all other text
      operators <- unique(c(raw_data()$Operator, raw_pipeline()$Operator))
      df <- full %>% select(-Operator, -Quarter_Order, -`MW Min`, -`MW Max`)   # Region kept for grouping

      if (nrow(df) == 0) {
        showNotification("No upcoming capacity in view, so there is nothing to export.",
                         type = "error", duration = 8)
        stop("No rows to export")
      }

      tryCatch(
        withProgress(message = "Building client PDF...", value = 0.5, {
          render_client_pdf(df, file, operators)
        }),
        error = function(e) {
          showNotification(paste("PDF export failed:", conditionMessage(e)),
                           type = "error", duration = 12)
          stop(e)
        }
      )
    },
    contentType = "application/pdf"
  )
  
  # ==========================================================
  # FIND CAPACITY
  #   Only sites with upcoming capacity are offered: a site with
  #   nothing in the pipeline is at capacity.
  #   Upcoming MW is summed per site up to the deadline (past
  #   quarters are ignored).
  #   MW range: minimum = must-have, ideal = target.
  #   Each site also carries a low/high estimate so a match can be
  #   labelled "firm" (even the low estimate meets the minimum).
  # ==========================================================
  
  fc_slider_max <- reactiveVal(100)
  sync_range("fc_need", fc_slider_max)
  
  observeEvent(pipeline_active(), {
    pl <- pipeline_active()
    
    updateSelectizeInput(
      session, "fc_state",
      choices = sort(unique(pl$State[pl$is_us & !is.na(pl$State) & nzchar(pl$State)])),
      server = TRUE
    )
    updateSelectizeInput(
      session, "fc_country",
      choices = sort(unique(na.omit(pl$Country))),
      server = TRUE
    )
    
    site_tot <- pl %>%
      group_by(Operator, City_clean, State, Country) %>%
      summarise(MW = sum(MW_available, na.rm = TRUE), .groups = "drop")
    
    mx <- suppressWarnings(max(site_tot$MW, na.rm = TRUE))
    if (!is.finite(mx)) mx <- 100
    mx <- max(ceiling(mx), 1)
    
    fc_slider_max(mx)
    updateSliderInput(session, "fc_need", max = mx,
                      value = c(min(10, mx), min(50, mx)))
  }, ignoreNULL = FALSE)
  
  fc_results <- reactive({
    rng <- input$fc_need
    req(length(rng) == 2, all(is.finite(rng)), input$fc_when)
    
    lo <- max(rng[1], 0.1)
    hi <- max(rng[2], lo)
    deadline <- quarter_to_date(input$fc_when)
    
    max_price <- input$fc_price
    max_pue <- input$fc_pue
    if (is.null(max_price) || !is.finite(max_price)) max_price <- NA_real_
    if (is.null(max_pue) || !is.finite(max_pue)) max_pue <- NA_real_
    
    site_key <- function(d) paste(d$Operator, d$City_clean, d$State, d$Country, sep = "|")
    
    # Price / PUE live on the current-site rows
    attrs <- raw_data() %>%
      mutate(key = site_key(.), Price = parse_price(Price_per_kW)) %>%
      group_by(key) %>%
      summarise(
        Price = Price[!is.na(Price)][1],
        PUE = PUE[!is.na(PUE)][1],
        A_Lat = Latitude[!is.na(Latitude)][1],
        A_Lng = Longitude[!is.na(Longitude)][1],
        .groups = "drop"
      )
    
    df <- pipeline_active() %>%
      mutate(key = site_key(.)) %>%
      filter(!is.na(Quarter_Date), Quarter_Date <= deadline, MW_available > 0)
    
    if (nrow(df) == 0) return(tibble())
    
    df <- df %>%
      arrange(key, Quarter_Date) %>%
      group_by(key, Operator, City_clean, State, Country) %>%
      mutate(cum = cumsum(MW_available)) %>%
      summarise(
        MW = sum(MW_available, na.rm = TRUE),
        MW_lo = sum(coalesce(MW_min, MW_available), na.rm = TRUE),
        MW_hi = sum(coalesce(MW_max, MW_available), na.rm = TRUE),
        Meets_in = Quarter[which(cum >= lo)[1]],
        Meets_date = Quarter_Date[which(cum >= lo)[1]],
        Last_Q = dplyr::last(Quarter),
        Last_date = dplyr::last(Quarter_Date),
        Lat = Latitude[!is.na(Latitude)][1],
        Lng = Longitude[!is.na(Longitude)][1],
        .groups = "drop"
      ) %>%
      left_join(attrs, by = "key") %>%
      transmute(
        Operator, Market = City_clean, State, Country, MW, MW_lo, MW_hi,
        Delivers = coalesce(Meets_in, Last_Q),
        Deliver_Date = coalesce(Meets_date, Last_date),
        Price, PUE,
        Latitude = coalesce(Lat, A_Lat), Longitude = coalesce(Lng, A_Lng)
      )
    
    # ---- US / Global scope ----
    if (identical(input$fc_scope, "us")) {
      df <- df %>% filter(Country == "United States")
      if (length(input$fc_state) > 0) df <- df %>% filter(State %in% input$fc_state)
    } else if (length(input$fc_country) > 0) {
      df <- df %>% filter(Country %in% input$fc_country)
    }
    
    if (nrow(df) == 0) return(df)
    
    df <- df %>%
      mutate(
        meets = MW >= lo,
        firm = meets & MW_lo >= lo,
        
        # 55 pts MW fit: 40 for reaching the minimum, up to 55 for reaching the ideal
        s_mw = case_when(
          MW >= hi ~ 55,
          MW >= lo ~ if (hi > lo) 40 + 15 * (MW - lo) / (hi - lo) else 55,
          TRUE ~ 40 * MW / lo
        ),
        # 15 pts timing: earlier than the deadline is better (full marks at 1 year early)
        s_time = if_else(
          meets,
          10 + 5 * pmin(pmax(as.numeric(deadline - Deliver_Date) / 365, 0), 1),
          4
        ),
        # 20 pts price
        s_price = case_when(
          is.na(max_price) ~ 20,
          is.na(Price) ~ 8,
          Price <= max_price ~ 20,
          TRUE ~ pmax(0, 20 * (1 - (Price - max_price) / max_price))
        ),
        # 10 pts PUE
        s_pue = case_when(
          is.na(max_pue) ~ 10,
          is.na(PUE) ~ 4,
          PUE <= max_pue ~ 10,
          TRUE ~ 0
        ),
        Score = round(s_mw + s_time + s_price + s_pue),
        
        span_txt = mw_span_txt(MW_lo, MW_hi),
        rng_txt = if_else(nzchar(span_txt), paste0(" [range ", span_txt, "]"), ""),
        
        mw_txt = case_when(
          MW >= hi ~ paste0("Delivers ", format(round(MW, 1), big.mark = ","),
                            " MW", rng_txt, " by ", Delivers, " (ideal ", hi, " met)"),
          meets ~ paste0("Delivers ", format(round(MW, 1), big.mark = ","),
                         " MW", rng_txt, " by ", Delivers, " (minimum ", lo, " met)"),
          TRUE ~ paste0("<span class='fc-gap'>", format(round(lo - MW, 1)),
                        " MW short of the ", lo, " MW minimum</span>", rng_txt)
        ),
        price_txt = case_when(
          is.na(max_price) ~ NA_character_,
          is.na(Price) ~ "<span class='fc-gap'>no price on file</span>",
          Price <= max_price ~ paste0("$", round(Price), "/kW within budget"),
          TRUE ~ paste0("<span class='fc-gap'>$", round(Price), "/kW over budget</span>")
        ),
        pue_txt = case_when(
          is.na(max_pue) ~ NA_character_,
          is.na(PUE) ~ "<span class='fc-gap'>no PUE on file</span>",
          PUE <= max_pue ~ paste0("PUE ", PUE, " OK"),
          TRUE ~ paste0("<span class='fc-gap'>PUE ", PUE, " too high</span>")
        ),
        Why = paste0(
          mw_txt,
          if_else(is.na(price_txt), "", paste0(" &middot; ", price_txt)),
          if_else(is.na(pue_txt), "", paste0(" &middot; ", pue_txt))
        )
      )
    
    if (!isTRUE(input$fc_near)) df <- df %>% filter(meets)
    
    switch(
      input$fc_sort,
      "Most MW" = df %>% arrange(desc(MW)),
      "Lowest price" = df %>% arrange(is.na(Price), Price, desc(Score)),
      df %>% arrange(desc(Score), desc(MW))
    )
  })
  
  output$fc_summary <- renderUI({
    df <- fc_results()
    rng <- input$fc_need
    
    if (nrow(df) == 0) {
      return(div(
        class = "fc-summary",
        "No sites with upcoming capacity match. Try a later date, a lower minimum, ",
        "a wider location, or turn on near misses."
      ))
    }
    
    full <- sum(df$meets)
    best <- df[1, ]
    
    div(
      class = "fc-summary",
      tags$b(full), paste0(" of ", nrow(df), " sites with upcoming capacity can deliver at least "),
      tags$b(rng[1]), paste0(" MW by ", input$fc_when, ". Top pick: "),
      tags$b(paste0(best$Operator, " \u2014 ", best$Market)),
      paste0(" (", best$Score, "% match).")
    )
  })
  
  output$fc_cards <- renderUI({
    df <- fc_results()
    if (nrow(df) == 0) return(NULL)
    
    top <- head(df, 3)
    esc <- function(x) htmltools::htmlEscape(coalesce(as.character(x), ""))
    
    div(class = "fc-grid", lapply(seq_len(nrow(top)), function(i) {
      r <- top[i, ]
      loc <- c(esc(r$Market), esc(r$State))
      loc <- loc[nzchar(loc)]
      
      div(
        class = paste("fc-card", if (i == 1) "rank-1"),
        div(class = "fc-rank", paste0("#", i, " MATCH")),
        div(class = "fc-name", r$Operator),
        div(class = "fc-sub", paste(loc, collapse = ", ")),
        div(
          class = "fc-score",
          div(class = "fc-bar", span(style = sprintf("width:%d%%", as.integer(r$Score)))),
          span(class = "fc-pct", paste0(r$Score, "%"))
        ),
        div(
          span(class = "fc-tag", paste("Delivers:", r$Delivers)),
          if (isTRUE(r$firm)) {
            span(class = "fc-tag fc-firm", "Firm: low estimate still meets minimum")
          } else if (isTRUE(r$meets)) {
            span(class = "fc-tag", "Depends on where the estimate lands")
          }
        ),
        div(class = "fc-why", style = "margin-top:8px", HTML(r$Why))
      )
    }))
  })
  
  output$fc_table <- renderDT({
    df <- fc_results()
    
    if (nrow(df) == 0) {
      return(datatable(
        tibble(Message = "No sites with upcoming capacity match this requirement."),
        rownames = FALSE, options = list(dom = "t")
      ))
    }
    
    out <- df %>%
      transmute(
        Operator, Market, State, Country,
        `MW by deadline` = round(MW, 1),
        `MW range` = if_else(nzchar(span_txt), span_txt, fmt_mw(MW)),
        Delivers,
        `$/kW` = round(Price), PUE, `Match %` = Score
      )
    
    datatable(
      out, rownames = FALSE, class = "compact hover", selection = "none",
      options = list(pageLength = 10, dom = "lfrtip")
    ) %>%
      add_color_bar(out, "Match %")
  })
  
  # Base map drawn once; results are swapped in through leafletProxy
  output$fc_map <- renderLeaflet({
    base_leaflet() %>%
      setView(lng = -40, lat = 30, zoom = 2)
  })
  
  # Remembers which sites were last shown, so the map only re-fits its
  # view when the set of sites changes (not on every slider nudge)
  fc_map_state <- new.env()
  fc_map_state$keys <- NULL
  
  observe({
    req(isTRUE(fc_map_ready()))
    
    df <- fc_results()
    if (nrow(df) > 0) df <- df %>% filter(!is.na(Latitude), !is.na(Longitude))
    
    proxy <- leafletProxy("fc_map") %>% clearMarkers()
    
    if (nrow(df) == 0) {
      fc_map_state$keys <- character(0)
      return()
    }
    
    pal <- colorNumeric(c("#DBEAFE", "#60A5FA", "#1D4ED8", "#0A1A4A"), domain = c(0, 100))
    mw_max <- max(df$MW, 1, na.rm = TRUE)
    
    df <- df %>%
      mutate(
        r_px = 6 + 10 * sqrt(pmin(MW / mw_max, 1)),
        mw_line = esc_html(mw_range_txt(MW, MW_lo, MW_hi)),
        pop = paste0(
          "<b>", esc_html(coalesce(Operator, "Unknown")), "</b><br>",
          esc_html(coalesce(Market, "")), "<br>",
          Score, "% match<br>", Why
        ),
        lab = paste0(
          esc_html(coalesce(Operator, "Unknown")), " \u2014 ",
          esc_html(coalesce(Market, "")), " \u00b7 ", mw_line, " by deadline"
        )
      )
    
    proxy %>%
      addCircleMarkers(
        data = df, lng = ~Longitude, lat = ~Latitude,
        radius = ~r_px, color = "#BFDBFE", weight = 1.5,
        fillColor = ~pal(Score), fillOpacity = 0.85,
        popup = ~pop,
        label = ~lapply(lab, htmltools::HTML)
      )
    
    keys <- sort(paste(df$Operator, df$Market, df$State, df$Country, sep = "|"))
    
    if (!identical(keys, fc_map_state$keys)) {
      proxy %>% fit_points(df$Latitude, df$Longitude)
      fc_map_state$keys <- keys
    }
  })
  
  # ==========================================================
  # REFRESH FROM DC.XLSX
  # ==========================================================
  
  output$refresh_status <- renderUI({
    refresh_trigger()
    
    versions <- list_versions()
    
    if (nrow(versions) == 0) {
      return(tags$p("No Excel refresh has been performed yet."))
    }
    
    latest <- versions %>% arrange(desc(version)) %>% slice(1)
    
    tags$p(
      class = "version-note",
      strong("Latest database update: "), latest$timestamp,
      tags$br(),
      "Version: ", latest$version,
      tags$br(),
      "Sites: ", latest$n_rows,
      tags$br(),
      "Upcoming entries: ", latest$n_pipeline_rows
    )
  })
  
  observeEvent(input$refresh_excel, {
    showModal(modalDialog(
      title = "Password required",
      passwordInput("refresh_pw", "Enter password to refresh from DC.xlsx"),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("confirm_refresh", "Refresh", class = "btn-primary")
      ),
      easyClose = TRUE
    ))
  })
  
  observeEvent(input$confirm_refresh, {
    if (!identical(input$refresh_pw, REFRESH_PASSWORD)) {
      showNotification("Incorrect password.", type = "error", duration = 5)
      return()
    }
    
    removeModal()
    
    if (!file.exists(MASTER_FILE)) {
      showNotification(
        paste0("Could not find ", MASTER_FILE,
               ". Make sure it is in the same folder as app.R."),
        type = "error", duration = 10
      )
      return()
    }
    
    tryCatch({
      withProgress(message = "Refreshing dashboard from DC.xlsx...", value = 0, {
        result <- import_master_excel(
          progress_fn = function(i, n, addr) {
            incProgress(0.9 / n, detail = paste("Geocoding:", addr))
          }
        )
      })
      
      raw_data(load_current())
      raw_pipeline(load_current_pipeline())
      refresh_trigger(refresh_trigger() + 1)
      
      showNotification(
        paste0("DC.xlsx imported: ", nrow(result$current), " sites and ",
               nrow(result$pipeline), " upcoming-capacity entries across ",
               length(result$quarters), " quarters."),
        type = "message", duration = 8
      )
      
      if (nrow(result$flags) > 0) {
        showNotification(
          paste0(nrow(result$flags), " capacity ",
                 if (nrow(result$flags) == 1) "entry needs" else "entries need",
                 " review - see Data quality check in Version History."),
          type = "warning", duration = 12
        )
      }
    }, error = function(e) {
      showNotification(
        paste("Excel refresh failed:", conditionMessage(e)),
        type = "error", duration = 12
      )
    })
  })
  
  # ==========================================================
  # DATA QUALITY CHECK
  # ==========================================================
  
  output$flag_summary <- renderUI({
    f <- import_flags()
    
    if (nrow(f) == 0) {
      return(div(
        class = "fc-summary dq-ok",
        icon("circle-check"),
        "Every capacity value in DC.xlsx was read cleanly."
      ))
    }
    
    div(
      class = "fc-summary",
      tags$b(nrow(f)), if (nrow(f) == 1) " entry" else " entries",
      " across ", tags$b(n_distinct(f$Excel_Row)),
      if (n_distinct(f$Excel_Row) == 1) " sheet row" else " sheet rows",
      " could not be read cleanly. The dashboard used the best reading shown below; ",
      "fix the text in DC.xlsx (or confirm it is fine) and refresh."
    )
  })
  
  output$flag_table <- renderDT({
    f <- import_flags()
    
    if (nrow(f) == 0) {
      return(datatable(
        tibble(Message = "Nothing to review."),
        rownames = FALSE, options = list(dom = "t")
      ))
    }
    
    f %>%
      select(
        `Excel row` = Excel_Row, Operator, Market, Field,
        `Value in sheet` = Value_in_sheet, `Read as` = Read_as, Issue
      ) %>%
      datatable(
        rownames = FALSE, class = "compact hover", selection = "none",
        options = list(pageLength = 10, dom = "lfrtip")
      )
  })
  
  output$download_flags <- downloadHandler(
    filename = function() {
      paste0("capacity_review_list_", format(Sys.Date(), "%Y%m%d"), ".csv")
    },
    content = function(file) {
      write.csv(import_flags(), file, row.names = FALSE, na = "")
    }
  )
  
  # ==========================================================
  # VERSION HISTORY
  # ==========================================================
  
  output$version_table <- renderDT({
    refresh_trigger()
    
    versions <- list_versions()
    
    if (nrow(versions) == 0) {
      return(datatable(tibble(Message = "No versions yet."), rownames = FALSE,
                       options = list(dom = "t")))
    }
    
    versions %>%
      arrange(desc(version)) %>%
      select(
        Version = version,
        Timestamp = timestamp,
        Note = note,
        Rows = n_rows,
        `Upcoming rows` = n_pipeline_rows
      ) %>%
      datatable(
        selection = "single",
        rownames = FALSE,
        class = "compact hover",
        options = list(pageLength = 10)
      )
  })
  
  observeEvent(input$version_table_rows_selected, {
    sel <- input$version_table_rows_selected
    req(sel)
    
    versions <- list_versions() %>% arrange(desc(version))
    v <- versions$version[sel]
    
    showModal(modalDialog(
      title = paste("Restore version", v, "?"),
      p(paste0(
        "This makes version ", v, " live again. ",
        "The current data is saved first, so you can undo this."
      )),
      passwordInput("restore_pw", "Enter password to restore"),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("confirm_restore", "Restore", class = "btn-danger")
      ),
      easyClose = TRUE
    ))
    
    session$userData$pending_restore <- v
  })
  
  observeEvent(input$confirm_restore, {
    if (!identical(input$restore_pw, REFRESH_PASSWORD)) {
      showNotification("Incorrect password.", type = "error", duration = 5)
      return()
    }
    
    v <- session$userData$pending_restore
    req(v)
    
    restore_version(v)
    
    raw_data(load_current())
    raw_pipeline(load_current_pipeline())
    refresh_trigger(refresh_trigger() + 1)
    
    removeModal()
    
    showNotification(paste("Restored version", v, "and set it live."), type = "message")
  })
}

# ============================================================
# RUN APP
# ============================================================

shinyApp(ui, server)
