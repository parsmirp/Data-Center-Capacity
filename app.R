# ============================================================
# Data Center Capacity Dashboard
#
# MASTER DATA:   DC.xlsx
# RUNTIME DATA:  dc_data.sqlite
#
# Workflow:
# 1. Edit & save DC.xlsx
# 2. Launch the app - it imports DC.xlsx automatically at startup
#    (only if the file changed since the last import)
# 3. (Optional) "Refresh from DC.xlsx" in Version History if you
#    edit DC.xlsx while the app is already running
#
# SQLite is the fast runtime database. Geocoding happens during
# import and reuses cached coordinates.
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

# Password required for "Refresh from DC.xlsx" and "Restore version".
REFRESH_PASSWORD <- Sys.getenv("DC_REFRESH_PASSWORD", unset = "bytebt")

# ---------- Quarter columns ----------

QUARTER_COLS <- c(
  "Q3 2026", "Q4 2026",
  "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027",
  "Q1 2028", "Q2 2028", "Q3 2028", "Q4 2028",
  "Q1 2029", "Q2 2029", "Q3 2029", "Q4 2029"
)

# ---------- Optional master-file columns ----------

MASTER_OPTIONAL_COLS <- c(
  "Market", "Address", "City", "State/Province", "Region", "Capacity Upload date",
  QUARTER_COLS,
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
# PARSING HELPERS
# ============================================================
# "20-40 MW" -> 30 | "3 MW + 6 MW" -> 9 | "18 MW" -> 18 | blank -> NA

parse_capacity_mw <- function(x) {
  x <- as.character(x)
  x <- str_replace_all(x, ",", "")
  
  vapply(x, function(val) {
    if (is.na(val)) return(NA_real_)
    
    parts <- str_split(val, "\\+")[[1]]
    part_vals <- vapply(parts, function(p) {
      nums <- str_extract_all(p, "[0-9]+\\.?[0-9]*")[[1]]
      if (length(nums) == 0) return(NA_real_)
      mean(as.numeric(nums))
    }, numeric(1))
    
    if (all(is.na(part_vals))) return(NA_real_)
    sum(part_vals, na.rm = TRUE)
  }, numeric(1), USE.NAMES = FALSE)
}

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
        Capacity_MW_est = double(), Utility_Rate = character(),
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
        MW_available = double(), Latitude = double(), Longitude = double()
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
  names(raw) <- str_trim(names(raw))
  
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
      Capacity_MW_est = parse_capacity_mw(`Total Capacity (MW)`),
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
  
  pipeline_quarters <- raw %>% select(all_of(QUARTER_COLS))
  
  pipeline <- bind_cols(pipeline_meta, pipeline_quarters) %>%
    pivot_longer(
      cols = all_of(QUARTER_COLS),
      names_to = "Quarter",
      values_to = "raw_val"
    ) %>%
    filter(!is.na(raw_val)) %>%
    mutate(
      MW_available = parse_capacity_mw(raw_val),
      Quarter_Date = quarter_to_date(Quarter)
    ) %>%
    filter(!is.na(MW_available)) %>%
    select(
      row_id, Operator, City_clean, State, Country, Region,
      is_us, Quarter, Quarter_Date, MW_available
    )
  
  list(current = current, pipeline = pipeline)
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
# PIPELINE COORDINATES
# ============================================================

attach_pipeline_coordinates <- function(pipeline_df, current_df) {
  if (nrow(pipeline_df) == 0) {
    pipeline_df$Latitude <- numeric(0)
    pipeline_df$Longitude <- numeric(0)
    return(pipeline_df)
  }
  
  if (!"Latitude" %in% names(pipeline_df)) pipeline_df$Latitude <- NA_real_
  if (!"Longitude" %in% names(pipeline_df)) pipeline_df$Longitude <- NA_real_
  
  coords <- current_df %>%
    select(City_clean, State, Country, Latitude, Longitude) %>%
    filter(!is.na(Latitude), !is.na(Longitude)) %>%
    distinct(City_clean, State, Country, .keep_all = TRUE) %>%
    rename(Current_Latitude = Latitude, Current_Longitude = Longitude)
  
  pipeline_df %>%
    left_join(coords, by = c("City_clean", "State", "Country")) %>%
    left_join(
      country_centroids %>%
        rename(Country_Latitude = Latitude, Country_Longitude = Longitude),
      by = "Country"
    ) %>%
    mutate(
      Latitude = coalesce(
        Latitude, Current_Latitude,
        if_else(str_ends(City_clean, fixed(" (no city provided)")), Country_Latitude, NA_real_)
      ),
      Longitude = coalesce(
        Longitude, Current_Longitude,
        if_else(str_ends(City_clean, fixed(" (no city provided)")), Country_Longitude, NA_real_)
      )
    ) %>%
    select(-Current_Latitude, -Current_Longitude, -Country_Latitude, -Country_Longitude)
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
    Utility_Rate = NA_character_, Price_per_kW = NA_character_,
    Cooling = NA_character_, PUE = NA_real_, Tax_Incentives = NA_character_,
    Deal_Reg = NA_character_, Notes = NA_character_, Contacts = NA_character_,
    City_clean = NA_character_, Latitude = NA_real_, Longitude = NA_real_
  )
  
  for (col in names(defaults)) {
    if (!col %in% names(df)) df[[col]] <- defaults[[col]]
  }
  
  df
}

load_current_pipeline <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_pipeline_current" %in% dbListTables(con)) {
    return(tibble(
      Operator = character(), City_clean = character(), State = character(),
      Country = character(), Region = character(), is_us = logical(),
      Quarter = character(), Quarter_Date = as.Date(character()),
      MW_available = double(), Latitude = double(), Longitude = double()
    ))
  }
  
  df <- dbReadTable(con, "dc_pipeline_current") %>% as_tibble()
  
  if ("Quarter_Date" %in% names(df)) df$Quarter_Date <- as.Date(df$Quarter_Date)
  df$is_us <- as.logical(df$is_us)
  
  df
}

list_versions <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_meta" %in% dbListTables(con)) return(tibble())
  
  dbReadTable(con, "dc_meta") %>% arrange(desc(version)) %>% as_tibble()
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

file_signature <- function() {
  unname(tools::md5sum(MASTER_FILE))
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
  names(raw) <- str_trim(names(raw))
  
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
  set_stored_signature(import_sig)
  
  list(current = parsed$current, pipeline = parsed$pipeline)
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
      "Startup import complete: %d sites, %d upcoming-capacity entries.\n",
      nrow(startup_import_result$current),
      nrow(startup_import_result$pipeline)
    ))
  }
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

    // Highlight button: off <-> on
   Shiny.addCustomMessageHandler('hl_state', function(on) {
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
    var v = /\s/.test(s.t) ? '"' + s.t + '"' : s.t;
    $inp().val(v).trigger('input');
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
.app-title img { height: 44px; border-radius: 8px; }
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
.kpi-amber  { --accent: #F59E0B; }

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

/* ---------- Misc ---------- */
.btn-outline-primary { --bs-btn-color: #93C5FD; --bs-btn-border-color: #2B4A7A; --bs-btn-hover-bg: var(--blue); --bs-btn-hover-border-color: var(--blue); }
.version-note { color: var(--ink-2); font-size: 13.5px; max-width: 70ch; }
:focus-visible { outline: 2px solid var(--blue-2); outline-offset: 2px; }
@media (max-width: 768px) {
  .app-title-sub { display: none; }
  .kpi-value { font-size: 24px; }
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
      condition = "input.main_tabs != 'Version History'",
      
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
        title = "Highlight sites with capacity arriving in the next 4 quarters"
      )
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
      div(
        class = "sb-presets",
        actionButton("q_2026", "2026", class = "btn-sm btn-outline-primary"),
        actionButton("q_2027", "2027", class = "btn-sm btn-outline-primary"),
        actionButton("q_2028", "2028", class = "btn-sm btn-outline-primary"),
        actionButton("q_2029", "2029", class = "btn-sm btn-outline-primary")
      ),
      selectizeInput("quarter_filter", NULL, choices = QUARTER_COLS, multiple = TRUE,
                     options = list(placeholder = "All quarters",
                                    plugins = list("remove_button"))),
      
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
                   "Bigger, darker dots = more MW. Click a dot for details."),
        leafletOutput("map", height = 520)
      ),
      
      layout_columns(
        col_widths = c(7, 5),
        card(
          card_title("ranking-star", "Top operators", "By estimated MW in view"),
          plotlyOutput("top_operators_chart", height = "330px")
        ),
        card(
          card_title("globe", "Capacity by country", "Share of MW in view"),
          plotlyOutput("country_share_chart", height = "330px")
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
        kpi("amber", "clock", "Upcoming entries in view", "vb_pipeline_entries"),
        kpi("blue", "building", "Sites in view", "vb_pipeline_sites")
      ),
      
      card(
        full_screen = TRUE,
        card_title("map-location-dot", "Upcoming capacity map",
                   "Each dot is one site and quarter"),
        leafletOutput("pipeline_map", height = 560)
      ),
      
      card(
        card_title("chart-column", "Upcoming capacity by quarter",
                   "Hover a bar to see the operator breakdown"),
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
          )
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

server <- function(input, output, session) {
  
  cur <- load_current()
  raw_data <- reactiveVal(cur)
  
  raw_pipeline <- reactiveVal(
    attach_pipeline_coordinates(load_current_pipeline(), cur)
  )
  
  cap_slider_max <- reactiveVal(100)
  pipe_slider_max <- reactiveVal(100)
  refresh_trigger <- reactiveVal(0)
  
  fmt <- function(x) format(round(x), big.mark = ",")
  
  # ----------------------------------------------------------
  # Header: "last updated" pill
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
    if (identical(input$main_tabs, "Upcoming Capacity")) {
      n <- nrow(filtered_pipeline())
      total <- nrow(raw_pipeline())
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
    df <- raw_pipeline()
    
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
  
  # Year shortcut buttons
  for (yr in 2026:2029) {
    local({
      y <- yr
      observeEvent(input[[paste0("q_", y)]], {
        updateSelectizeInput(
          session, "quarter_filter",
          selected = QUARTER_COLS[grepl(as.character(y), QUARTER_COLS, fixed = TRUE)]
        )
      })
    })
  }
  
  # ----------------------------------------------------------
  # "Coming soon" highlight logic
  # ----------------------------------------------------------
  
  upcoming_soon_keys <- reactive({
    pl <- raw_pipeline()
    
    if (nrow(pl) == 0) return(character(0))
    
    today_q <- as.Date(cut(Sys.Date(), "quarter"))
    
    soon_quarters <- pl %>%
      filter(Quarter_Date >= today_q) %>%
      distinct(Quarter, Quarter_Date) %>%
      arrange(Quarter_Date) %>%
      slice_head(n = 4) %>%
      pull(Quarter)
    
    pl %>%
      filter(Quarter %in% soon_quarters) %>%
      mutate(key = paste(Operator, City_clean, State, Country, sep = "|")) %>%
      pull(key) %>%
      unique()
  })
  
  # Count of sites in view that are coming online in the next 4 quarters
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
  # DATA CENTERS MAP
  # ==========================================================
  
  output$map <- renderLeaflet({
    leaflet(options = leafletOptions(worldCopyJump = FALSE, minZoom = 2, maxZoom = 18)) %>%
      addTiles(
        urlTemplate = CARTO_POSITRON_URL,
        attribution = CARTO_ATTRIBUTION,
        options = tileOptions(noWrap = TRUE)
      ) %>%
      setMaxBounds(lng1 = -180, lat1 = -85, lng2 = 180, lat2 = 85) %>%
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
  # orange = arriving in the next 4 quarters (when highlight is on)
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
        
        op_txt = htmltools::htmlEscape(enc2utf8(coalesce(Operator, "Unknown operator"))),
        city_txt = htmltools::htmlEscape(enc2utf8(coalesce(City_clean, ""))),
        place_txt = htmltools::htmlEscape(enc2utf8(
          if_else(is.na(State) | State == "", coalesce(Country, ""), State)
        )),
        cap_txt = htmltools::htmlEscape(enc2utf8(coalesce(Capacity, "No info"))),
        cool_txt = htmltools::htmlEscape(enc2utf8(coalesce(Cooling, "Unknown"))),
        notes_txt = htmltools::htmlEscape(enc2utf8(coalesce(Notes, ""))),
        popup_html = paste0(
          "<div style='min-width:200px; max-width:280px'>",
          "<b>", op_txt, "</b><br>",
          city_txt, ", ", place_txt, "<br>",
          "Capacity: ", cap_txt,
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
        label_txt = paste0(op_txt, " \u2014 ", city_txt)
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
          6 + 12 * sqrt((Capacity_MW_est - min_mw) / (max_mw - min_mw)),
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
        "Coming online in next 4 quarters</div>"
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
      mutate(Operator = coalesce(Operator, "Unknown operator")) %>%
      group_by(Operator) %>%
      summarise(MW = sum(Capacity_MW_est, na.rm = TRUE),
                Sites = n(), .groups = "drop") %>%
      arrange(desc(MW)) %>%
      slice_head(n = 10)
    
    if (nrow(df) == 0) return(empty_plot("No capacity matches the current search and filters."))
    
    df <- df %>% mutate(Operator = factor(Operator, levels = rev(Operator)))
    
    plot_ly(
      df, x = ~MW, y = ~Operator, type = "bar", orientation = "h",
      text = ~fmt(MW), textposition = "outside", cliponaxis = FALSE,
      textfont = list(color = "#CBD5E1"),
      hovertext = ~paste0("<b>", Operator, "</b><br>", fmt(MW), " MW across ",
                          Sites, if_else(Sites == 1L, " site", " sites")),
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
  
  output$country_share_chart <- renderPlotly({
    df <- filtered() %>%
      filter(!is.na(Capacity_MW_est), Capacity_MW_est > 0) %>%
      mutate(Country = coalesce(Country, "Unknown")) %>%
      group_by(Country) %>%
      summarise(MW = sum(Capacity_MW_est, na.rm = TRUE), .groups = "drop") %>%
      arrange(desc(MW))
    
    if (nrow(df) == 0) return(empty_plot("No capacity matches the current search and filters."))
    
    if (nrow(df) > 7) {
      df <- bind_rows(
        df %>% slice_head(n = 7),
        tibble(Country = "Other", MW = sum(df$MW[8:nrow(df)]))
      )
    }
    
    cols <- c("#1D4ED8", "#3B82F6", "#60A5FA", "#93C5FD", "#2DD4BF",
              "#F59E0B", "#A78BFA", "#64748B")[seq_len(nrow(df))]
    
    plot_ly(
      df, labels = ~Country, values = ~MW, type = "pie", hole = 0.62,
      sort = FALSE, direction = "clockwise",
      textinfo = "none",
      hovertemplate = "<b>%{label}</b><br>%{value:,.0f} MW (%{percent})<extra></extra>",
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
  # ==========================================================
  
  output$pipeline_map <- renderLeaflet({
    df <- filtered_pipeline()
    searching <- nzchar(search_raw_d())
    
    base_map <- leaflet(options = leafletOptions(worldCopyJump = FALSE, minZoom = 2, maxZoom = 18)) %>%
      addTiles(
        urlTemplate = CARTO_POSITRON_URL,
        attribution = CARTO_ATTRIBUTION,
        options = tileOptions(noWrap = TRUE)
      ) %>%
      setMaxBounds(lng1 = -180, lat1 = -85, lng2 = 180, lat2 = 85) %>%
      setView(
        lng = if (identical(input$pipeline_scope, "us")) -98.5 else -20,
        lat = if (identical(input$pipeline_scope, "us")) 39.5 else 25,
        zoom = if (identical(input$pipeline_scope, "us")) 4 else 2
      ) %>%
      addEasyButton(reset_view_button("pipeline_map_reset"))
    
    if (nrow(df) == 0) return(base_map)
    
    df <- df %>%
      left_join(
        country_centroids %>%
          rename(Country_Latitude = Latitude, Country_Longitude = Longitude),
        by = "Country"
      ) %>%
      mutate(
        map_base_lat = coalesce(Latitude, Country_Latitude),
        map_base_lng = coalesce(Longitude, Country_Longitude)
      )
    
    # Spread rows at the same location so each one can be clicked
    df <- df %>%
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
      ungroup()
    
    map_df <- df %>%
      filter(!is.na(map_lat), !is.na(map_lng), !is.na(MW_available), MW_available >= 0)
    
    if (nrow(map_df) == 0) return(base_map)
    
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
    
    esc <- function(x) htmltools::htmlEscape(enc2utf8(as.character(x)))
    
    popup_text <- paste0(
      "<div style='min-width:210px'>",
      "<b style='font-size:15px'>",
      ifelse(is.na(map_df$Operator), "Unknown operator", esc(map_df$Operator)),
      "</b><br>",
      ifelse(
        is.na(map_df$City_clean) | map_df$City_clean == "",
        esc(map_df$Country),
        esc(map_df$City_clean)
      ),
      ifelse(
        !is.na(map_df$State) & map_df$State != "",
        paste0(", ", esc(map_df$State)),
        ""
      ),
      "<br><br>",
      "<b>Coming online:</b> ", map_df$Quarter, "<br>",
      "<b>MW available:</b> ",
      format(round(map_df$MW_available, 1), big.mark = ","), " MW",
      "</div>"
    )
    
    m <- base_map %>%
      addCircleMarkers(
        lng = map_df$map_lng,
        lat = map_df$map_lat,
        radius = radius,
        stroke = TRUE,
        weight = 1.5,
        color = "#BFDBFE",
        fillColor = pal(map_df$MW_available),
        fillOpacity = 0.85,
        popup = popup_text
      ) %>%
      addLegend(
        position = "bottomright",
        pal = pal,
        values = map_df$MW_available,
        title = "Upcoming MW",
        opacity = 0.9,
        labFormat = labelFormat(suffix = " MW")
      )
    
    if (searching) m <- fit_points(m, map_df$map_lat, map_df$map_lng)
    
    m
  })
  
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
        Quarter = factor(Quarter, levels = QUARTER_COLS),
        Operator = if_else(is.na(Operator) | Operator == "", "Unknown operator", Operator)
      )
    
    quarter_totals <- df %>%
      group_by(Quarter) %>%
      summarise(MW = sum(MW_available, na.rm = TRUE), .groups = "drop")
    
    operator_breakdown <- df %>%
      group_by(Quarter, Operator) %>%
      summarise(MW = sum(MW_available, na.rm = TRUE), .groups = "drop") %>%
      arrange(Quarter, desc(MW)) %>%
      group_by(Quarter) %>%
      summarise(
        Breakdown = paste0(
          "<b>", Operator, "</b>: ",
          format(round(MW, 1), big.mark = ","), " MW",
          collapse = "<br>"
        ),
        .groups = "drop"
      )
    
    chart_data <- quarter_totals %>% left_join(operator_breakdown, by = "Quarter")
    
    plot_ly(
      data = chart_data,
      x = ~Quarter, y = ~MW, type = "bar",
      text = ~format(round(MW, 1), big.mark = ","),
      textposition = "outside", cliponaxis = FALSE,
      textfont = list(color = "#CBD5E1"),
      hovertext = ~paste0(
        "<b>", Quarter, "</b>",
        "<br><b>Total: ", format(round(MW, 1), big.mark = ","), " MW</b>",
        "<br><br>", Breakdown
      ),
      hoverinfo = "text",
      marker = list(color = "#3B82F6")
    ) %>%
      dark_plot(
        xaxis = list(
          title = NULL, categoryorder = "array", categoryarray = QUARTER_COLS,
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
      mutate(Quarter_Order = match(Quarter, QUARTER_COLS)) %>%
      arrange(Quarter_Order, Operator, Country, State, City_clean) %>%
      transmute(
        Operator,
        City = City_clean,
        State,
        Country,
        Region,
        Quarter,
        `MW Available` = MW_available,
        Quarter_Order
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
        order = list(list(7, "asc")),
        columnDefs = list(list(visible = FALSE, targets = 7))
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
      raw_pipeline(attach_pipeline_coordinates(load_current_pipeline(), load_current()))
      refresh_trigger(refresh_trigger() + 1)
      
      showNotification(
        paste0("DC.xlsx imported: ", nrow(result$current), " sites and ",
               nrow(result$pipeline), " upcoming-capacity entries."),
        type = "message", duration = 8
      )
    }, error = function(e) {
      showNotification(
        paste("Excel refresh failed:", conditionMessage(e)),
        type = "error", duration = 12
      )
    })
  })
  
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
    raw_pipeline(attach_pipeline_coordinates(load_current_pipeline(), load_current()))
    refresh_trigger(refresh_trigger() + 1)
    
    removeModal()
    
    showNotification(paste("Restored version", v, "and set it live."), type = "message")
  })
}

# ============================================================
# RUN APP
# ============================================================

shinyApp(ui, server)
