# ============================================================
# Data Center Capacity Dashboard
# MASTER DATA:
# DC.xlsx (THANOS EXCEL SHEET)


# RUNTIME DATA:
# dc_data.sqlite

# Workflow:
# 1. Edit & save DC.xlsx
# 2. Launch the app - it imports DC.xlsx automatically at startup
#    (only if the file changed since the last import)
# 3. (Optional) Use "Refresh from DC.xlsx" in Version History if
#    you edit DC.xlsx while the app is already running

# SQLite remains the fast runtime database.
# Geocoding happens during import (startup or manual refresh),
# reusing cached coordinates for cities already geocoded.
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

# ---------- File locations ----------

DB_PATH <- "dc_data.sqlite"
MASTER_FILE <- "DC.xlsx"

## carto map api key
CARTO_KEY <- Sys.getenv("CARTO_API_KEY")
if (!nzchar(CARTO_KEY) && file.exists("carto_key.txt")) {
  CARTO_KEY <- trimws(readLines("carto_key.txt", n=1, warn = FALSE))
}
CARTO_POSITRON_URL <- paste0(
  "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
  if (nzchar(CARTO_KEY)) paste0("?key=", CARTO_KEY) else ""
)
CARTO_ATTRIBUTION <- paste0(
  '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> ',
  'contributors &copy; <a href="https://carto.com/attributions">CARTO</a>'
)
##

REFRESH_PASSWORD <- Sys.getenv("DC_REFRESH_PASSWORD", unset = "bytebt")
# Password required to click "Refresh from DC.xlsx".
# Reads the DC_REFRESH_PASSWORD environment variable if set,
# otherwise falls back to the value below.

# ---------- Quarter columns ----------

QUARTER_COLS <- c(
  "Q3 2026", "Q4 2026",
  "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027",
  "Q1 2028", "Q2 2028", "Q3 2028", "Q4 2028",
  "Q1 2029", "Q2 2029", "Q3 2029", "Q4 2029"
)

# ---------- Optional master-file columns ----------

MASTER_OPTIONAL_COLS <- c(
  "City",
  "State/Province",
  "Region",
  "Capacity Upload date",
  QUARTER_COLS,
  "Utility Rate ($/kWh)",
  "Price ($/kW)",
  "Cooling",
  "PUE",
  "Tax Incentives",
  "Deal Reg",
  "Notes",
  "Contacts"
)

# ---------- Country centroids ----------
# Used only when a row has no city.
# These are approximate map locations, not exact site locations.

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

# ============================================================
# PARSING HELPERS
# ============================================================
# Examples:
# "20-40 MW"    -> 30
# "3 MW + 6 MW" -> 9
# "18 MW"       -> 18
# "5-10 MW ramp"-> 7.5
# NA / blank    -> NA

parse_capacity_mw <- function(x) {
  x <- as.character(x)
  
  # Remove thousands-separator commas (e.g. "3,300" -> "3300")
  # before any number extraction happens below.
  x <- str_replace_all(x, ",", "")
  
  vapply(x, function(val) {
    if (is.na(val)) {
      return(NA_real_)
    }
    
    parts <- str_split(val, "\\+")[[1]]
    part_vals <- vapply(parts, function(p) {
      nums <- str_extract_all(p, "[0-9]+\\.?[0-9]*")[[1]]
      
      if (length(nums) == 0) {
        return(NA_real_)
      }
      
      mean(as.numeric(nums))
    }, numeric(1))
    
    if (all(is.na(part_vals))) {
      return(NA_real_)
    }
    
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
# DATABASE HELPERS
# ============================================================

get_con <- function() {
  dbConnect(RSQLite::SQLite(), DB_PATH)
}

initialize_database <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  # Create metadata table if necessary.
  if (!"dc_meta" %in% dbListTables(con)) {
    dbWriteTable(
      con,
      "dc_meta",
      tibble(
        version = integer(),
        timestamp = character(),
        note = character(),
        n_rows = integer(),
        n_pipeline_rows = integer()
      ),
      overwrite = TRUE
    )
  }
  
  # Create an empty current table if necessary.
  if (!"dc_current" %in% dbListTables(con)) {
    dbWriteTable(
      con,
      "dc_current",
      tibble(
        Operator = character(),
        City = character(),
        State = character(),
        Country = character(),
        Region = character(),
        Capacity = character(),
        Capacity_MW_est = double(),
        Utility_Rate = character(),
        Price_per_kW = character(),
        Cooling = character(),
        PUE = double(),
        Tax_Incentives = character(),
        Deal_Reg = character(),
        Notes = character(),
        Contacts = character(),
        is_us = logical(),
        City_clean = character(),
        Latitude = double(),
        Longitude = double()
      ),
      overwrite = TRUE
    )
  }
  
  # Create empty pipeline table.
  if (!"dc_pipeline_current" %in% dbListTables(con)) {
    dbWriteTable(
      con,
      "dc_pipeline_current",
      tibble(
        Operator = character(),
        City_clean = character(),
        State = character(),
        Country = character(),
        Region = character(),
        is_us = logical(),
        Quarter = character(),
        Quarter_Date = character(),
        MW_available = double(),
        Latitude = double(),
        Longitude = double()
      ),
      overwrite = TRUE
    )
  }
  
  # Geocode cache.
  if (!"geocode_cache" %in% dbListTables(con)) {
    dbWriteTable(
      con,
      "geocode_cache",
      tibble(
        key = character(),
        Latitude = double(),
        Longitude = double()
      ),
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
  
  # Add optional columns if missing.
  for (col in MASTER_OPTIONAL_COLS) {
    if (!col %in% names(raw)) {
      raw[[col]] <- NA_character_
    }
  }
  
  # Trim text and turn blank strings into NA.
  raw <- raw %>%
    mutate(across(everything(), ~ na_if(str_trim(as.character(.x)), "")))
  
  # ---------- Current sites ----------
  current <- raw %>%
    transmute(
      Operator = `DC Operator`,
      City = City,
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
      City_clean = if_else(!is.na(City), City, paste0(Country, " (no city provided)"))
    )
  
  # ---------- Pipeline metadata ----------
  pipeline_meta <- raw %>%
    transmute(
      Operator = `DC Operator`,
      City = City,
      State = `State/Province`,
      Country = Country,
      Region = Region,
      is_us = Country == "United States",
      City_clean = if_else(!is.na(City), City, paste0(Country, " (no city provided)"))
    )
  
  pipeline_quarters <- raw %>%
    select(all_of(QUARTER_COLS))
  
  # ---------- Convert quarter columns into rows ----------
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
      Operator, City_clean, State, Country, Region,
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
    return(tibble(
      key = character(),
      Latitude = double(),
      Longitude = double()
    ))
  }
  
  dbReadTable(con, "geocode_cache") %>%
    as_tibble()
}

save_geocode_cache <- function(cache_df) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  dbWriteTable(con, "geocode_cache", cache_df, overwrite = TRUE)
}

# Geocode during import (startup or manual refresh).
# Existing cached cities are instantaneous.
# New cities use OSM if tidygeocoder is installed.

geocode_current <- function(current_df,
                            existing_current = NULL,
                            progress_fn = NULL) {
  cache <- load_geocode_cache()
  
  # Preserve coordinates already stored in SQLite.
  if (!is.null(existing_current) &&
      all(c("City_clean", "State", "Country", "Latitude", "Longitude") %in%
          names(existing_current))) {
    existing_coords <- existing_current %>%
      filter(!is.na(Latitude), !is.na(Longitude)) %>%
      transmute(
        key = paste(City_clean, State, Country, sep = "|"),
        Latitude = as.numeric(Latitude),
        Longitude = as.numeric(Longitude)
      ) %>%
      distinct(key, .keep_all = TRUE)
    
    cache <- bind_rows(cache, existing_coords) %>%
      distinct(key, .keep_all = TRUE)
  }
  
  # Find cities not already cached.
  to_geocode <- current_df %>%
    filter(!is.na(City)) %>%
    distinct(City_clean, State, Country) %>%
    mutate(key = paste(City_clean, State, Country, sep = "|")) %>%
    filter(!key %in% cache$key)
  
  new_rows <- tibble()
  
  # Geocode new cities.
  if (nrow(to_geocode) > 0 &&
      requireNamespace("tidygeocoder", quietly = TRUE)) {
    for (i in seq_len(nrow(to_geocode))) {
      r <- to_geocode[i, ]
      
      addr <- paste(na.omit(c(r$City_clean, r$State, r$Country)), collapse = ", ")
      
      if (!is.null(progress_fn)) {
        progress_fn(i, nrow(to_geocode), addr)
      }
      
      res <- tryCatch(
        tidygeocoder::geo(address = addr, method = "osm", quiet = TRUE),
        error = function(e) NULL
      )
      
      if (!is.null(res) &&
          nrow(res) > 0 &&
          !is.na(res$lat[1]) &&
          !is.na(res$long[1])) {
        new_rows <- bind_rows(
          new_rows,
          tibble(
            key = r$key,
            Latitude = as.numeric(res$lat[1]),
            Longitude = as.numeric(res$long[1])
          )
        )
      }
      
      # Nominatim rate limit.
      Sys.sleep(1)
    }
  }
  
  # Save newly geocoded cities.
  if (nrow(new_rows) > 0) {
    cache <- bind_rows(cache, new_rows) %>%
      distinct(key, .keep_all = TRUE)
    
    save_geocode_cache(cache)
  }
  
  # Apply coordinates.
  current_df %>%
    mutate(key = paste(City_clean, State, Country, sep = "|")) %>%
    left_join(cache %>% select(key, Latitude, Longitude), by = "key") %>%
    left_join(
      country_centroids %>%
        rename(Country_Lat = Latitude, Country_Lon = Longitude),
      by = "Country"
    ) %>%
    mutate(
      Latitude = case_when(
        !is.na(Latitude) ~ Latitude,
        is.na(City) & !is.na(Country_Lat) ~ Country_Lat,
        TRUE ~ NA_real_
      ),
      Longitude = case_when(
        !is.na(Longitude) ~ Longitude,
        is.na(City) & !is.na(Country_Lon) ~ Country_Lon,
        TRUE ~ NA_real_
      )
    ) %>%
    select(-key, -Country_Lat, -Country_Lon)
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
  
  if (!"Latitude" %in% names(pipeline_df)) {
    pipeline_df$Latitude <- NA_real_
  }
  
  if (!"Longitude" %in% names(pipeline_df)) {
    pipeline_df$Longitude <- NA_real_
  }
  
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
        Latitude,
        Current_Latitude,
        if_else(
          str_ends(City_clean, fixed(" (no city provided)")),
          Country_Latitude,
          NA_real_
        )
      ),
      Longitude = coalesce(
        Longitude,
        Current_Longitude,
        if_else(
          str_ends(City_clean, fixed(" (no city provided)")),
          Country_Longitude,
          NA_real_
        )
      )
    ) %>%
    select(
      -Current_Latitude, -Current_Longitude,
      -Country_Latitude, -Country_Longitude
    )
}

# ============================================================
# DATABASE READ FUNCTIONS
# ============================================================

load_current <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  df <- dbReadTable(con, "dc_current") %>%
    as_tibble()
  
  if ("is_us" %in% names(df)) {
    df$is_us <- as.logical(df$is_us)
  } else {
    df$is_us <- FALSE
  }
  
  # Make sure expected columns exist.
  defaults <- list(
    Operator = NA_character_,
    City = NA_character_,
    State = NA_character_,
    Country = NA_character_,
    Region = NA_character_,
    Capacity = NA_character_,
    Capacity_MW_est = NA_real_,
    Utility_Rate = NA_character_,
    Price_per_kW = NA_character_,
    Cooling = NA_character_,
    PUE = NA_real_,
    Tax_Incentives = NA_character_,
    Deal_Reg = NA_character_,
    Notes = NA_character_,
    Contacts = NA_character_,
    City_clean = NA_character_,
    Latitude = NA_real_,
    Longitude = NA_real_
  )
  
  for (col in names(defaults)) {
    if (!col %in% names(df)) {
      df[[col]] <- defaults[[col]]
    }
  }
  
  df
}

load_current_pipeline <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_pipeline_current" %in% dbListTables(con)) {
    return(
      tibble(
        Operator = character(),
        City_clean = character(),
        State = character(),
        Country = character(),
        Region = character(),
        is_us = logical(),
        Quarter = character(),
        Quarter_Date = as.Date(character()),
        MW_available = double(),
        Latitude = double(),
        Longitude = double()
      )
    )
  }
  
  df <- dbReadTable(con, "dc_pipeline_current") %>%
    as_tibble()
  
  if ("Quarter_Date" %in% names(df)) {
    df$Quarter_Date <- as.Date(df$Quarter_Date)
  }
  
  df$is_us <- as.logical(df$is_us)
  
  df
}

list_versions <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"dc_meta" %in% dbListTables(con)) {
    return(tibble())
  }
  
  dbReadTable(con, "dc_meta") %>%
    arrange(desc(version)) %>%
    as_tibble()
}

# ============================================================
# SAVE VERSION
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
  
  # Current metadata.
  meta <- if ("dc_meta" %in% dbListTables(con)) {
    dbReadTable(con, "dc_meta") %>%
      as_tibble()
  } else {
    tibble(
      version = integer(),
      timestamp = character(),
      note = character(),
      n_rows = integer(),
      n_pipeline_rows = integer()
    )
  }
  
  if (!"n_pipeline_rows" %in% names(meta)) {
    meta$n_pipeline_rows <- NA_integer_
  }
  
  next_v <- if (nrow(meta) == 0) {
    1L
  } else {
    max(meta$version) + 1L
  }
  
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
  
  # Write new live data.
  dbWriteTable(con, "dc_current", new_current, overwrite = TRUE)
  dbWriteTable(con, "dc_pipeline_current", new_pipeline, overwrite = TRUE)
  
  # Save the new version.
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

# ============================================================
# RESTORE VERSION
# ============================================================

restore_version <- function(v) {
  con <- get_con()
  
  tbl_name <- paste0("dc_history_v", v)
  
  if (!tbl_name %in% dbListTables(con)) {
    dbDisconnect(con)
    return(FALSE)
  }
  
  old_current <- dbReadTable(con, tbl_name)
  
  pipe_tbl <- paste0("dc_pipeline_history_v", v)
  
  old_pipeline <- if (pipe_tbl %in% dbListTables(con)) {
    dbReadTable(con, pipe_tbl)
  } else {
    tibble()
  }
  
  dbDisconnect(con)
  
  save_new_version(old_current, old_pipeline, note = paste("Restored from version", v))
  
  TRUE
}

# ============================================================
# IMPORT CHANGE DETECTION
# Skips the startup import when DC.xlsx hasn't changed
# ============================================================

file_signature <- function() {
  unname(tools::md5sum(MASTER_FILE))
}

get_stored_signature <- function() {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  if (!"import_state" %in% dbListTables(con)) {
    return(NA_character_)
  }
  
  df <- dbReadTable(con, "import_state")
  
  if (nrow(df) == 0) NA_character_ else df$signature[1]
}

set_stored_signature <- function(sig) {
  con <- get_con()
  on.exit(dbDisconnect(con), add = TRUE)
  
  dbWriteTable(
    con,
    "import_state",
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
  
  # Hash the file before reading it, so an edit made mid-import
  # isn't mistaken for "already imported".
  import_sig <- file_signature()
  
  # Read Excel.
  raw <- readxl::read_excel(MASTER_FILE, col_types = "text")
  names(raw) <- str_trim(names(raw))
  
  # Clean into current + pipeline.
  parsed <- clean_master_df(raw)
  
  # Existing data.
  existing_current <- tryCatch(load_current(), error = function(e) NULL)
  
  # Geocode / preserve coordinates.
  parsed$current <- geocode_current(
    parsed$current,
    existing_current = existing_current,
    progress_fn = progress_fn
  )
  
  # Geocode the unique locations represented in the upcoming-capacity
  # pipeline as well. This means a pipeline-only location can still
  # appear on the map even if it is not yet a current site.
  pipeline_locations <- parsed$pipeline %>%
    distinct(City_clean, State, Country) %>%
    mutate(
      City = if_else(
        str_ends(City_clean, fixed(" (no city provided)")),
        NA_character_,
        City_clean
      ),
      Operator = NA_character_,
      Region = NA_character_,
      is_us = Country == "United States"
    )
  
  if (nrow(pipeline_locations) > 0) {
    pipeline_locations <- geocode_current(
      pipeline_locations,
      existing_current = parsed$current,
      progress_fn = progress_fn
    )
  }
  
  parsed$pipeline <- parsed$pipeline %>%
    left_join(
      pipeline_locations %>%
        select(City_clean, State, Country, Latitude, Longitude) %>%
        distinct(City_clean, State, Country, .keep_all = TRUE),
      by = c("City_clean", "State", "Country")
    )
  
  # Save to SQLite.
  save_new_version(parsed$current, parsed$pipeline, note = "Refresh from DC.xlsx")
  set_stored_signature(import_sig)
  
  list(current = parsed$current, pipeline = parsed$pipeline)
}

# ============================================================
# STARTUP IMPORT FROM DC.XLSX
#
# Runs once when the app process starts, but only if DC.xlsx
# changed since the last import. import_master_excel() calls
# save_new_version(), which snapshots the previous state first,
# so even a startup import is reversible from the Version
# History tab.
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
    cat(
      sprintf(
        "Startup import complete: %d sites, %d upcoming-capacity entries.\n",
        nrow(startup_import_result$current),
        nrow(startup_import_result$pipeline)
      )
    )
  }
}

# ============================================================
# JS + small UI helpers
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

  // Highlight button: switch between outline and filled style
  Shiny.addCustomMessageHandler('hl_state', function(on) {
    $('#highlight_toggle')
      .toggleClass('btn-warning', on)
      .toggleClass('btn-outline-warning', !on)
      .find('.hl-label')
      .text(on ? 'Highlighting coming soon' : 'Highlight coming soon');
  });
});
)---"

# A range slider with two editable number boxes on top.
# Drag the slider OR type exact values in the boxes; they stay in sync
# (see sync_range() in the server).
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

# ============================================================
# UI
# ============================================================

ui <- page_sidebar(
  title = tags$span(
    tags$img(
      src = "logo.png",
      height = "80px",
      style = "margin-right:10px; vertical-align:middle;",
      onerror = "this.style.display='none'"
    )
  ),
  
  theme = bs_theme(
    version = 5,
    bg = "#0B0F14",
    fg = "#E2E8F0",
    primary = "#3B82F6",
    secondary = "#64748B",
    success = "#22C55E",
    warning = "#F59E0B",
    danger = "#EF4444",
    base_font = font_google("Inter", local = TRUE),
    heading_font = font_google("Inter", local = TRUE)
  ) %>%
    bs_add_rules(
      "
  html, body,
  .bslib-page-sidebar,
  .bslib-sidebar-layout,
  .bslib-sidebar-layout > .main,
  .tab-content,
  .container-fluid {
    background-color: #0B0F14 !important;
    color: #E2E8F0 !important;
  }

  .card,
  .card-header,
  .card-body {
    background-color: #111827 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  .bslib-value-box,
  .bslib-value-box .value-box-area,
  .bslib-value-box .value-box-showcase {
    background-color: #111827 !important;
    border-color: #1F2937 !important;
    color: #E2E8F0 !important;
  }

  .bslib-sidebar-layout > .sidebar {
    background-color: #111827 !important;
    color: #E2E8F0 !important;
  }

  .leaflet-container {
    background-color: #0B0F14 !important;
  }

  .dataTables_wrapper {
    background-color: #111827 !important;
    color: #E2E8F0 !important;
  }

  table.dataTable,
  table.dataTable td,
  table.dataTable th {
    background-color: #111827 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  table.dataTable tbody tr:hover {
    background-color: #1F2937 !important;
  }

  .dataTables_wrapper .dataTables_length,
  .dataTables_wrapper .dataTables_filter,
  .dataTables_wrapper .dataTables_info,
  .dataTables_wrapper .dataTables_paginate {
    color: #E2E8F0 !important;
  }

  .dataTables_wrapper .form-control {
    background-color: #0B0F14 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  #pipeline_mw_filter {
    background-color: #0B0F14 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  .form-control,
  .selectize-input {
    background-color: #0B0F14 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  .selectize-dropdown {
    background-color: #111827 !important;
    color: #E2E8F0 !important;
    border-color: #1F2937 !important;
  }

  /* Make the lightning bolt icon yellow and glowing */
  .bslib-value-box .value-box-showcase .fa-bolt,
  .bslib-value-box .value-box-showcase .fas.fa-bolt {
    color: #FDE047 !important;
    text-shadow:
      0 0 6px  rgba(250, 204, 21, 0.7),
      0 0 12px rgba(234, 179, 8, 0.4);
  }

  /* Subtle theme-blue glow for Operators / Cities icons */
  .bslib-value-box .value-box-showcase .glow-blue i,
  .bslib-value-box .value-box-showcase .glow-blue svg {
    color: #3B82F6 !important;
    fill: #3B82F6 !important;
    text-shadow: 0 0 6px rgba(59, 130, 246, 0.6);
    filter: drop-shadow(0 0 3px rgba(59, 130, 246, 0.5));
  }

  /* ---------- Sidebar polish ---------- */
  .sb-summary {
    background: linear-gradient(135deg, rgba(59,130,246,.18), rgba(59,130,246,.04));
    border: 1px solid #1F2937; border-radius: 10px;
    padding: 10px 12px; margin-bottom: 12px;
  }
  .sb-summary-main { font-size: 13px; color: #CBD5E1; }
  .sb-summary-main b { font-size: 22px; color: #60A5FA; }

  .sb-presets { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 12px; }

  .sb-chips { margin-bottom: 8px; }
  .sb-chip {
    display: inline-flex; align-items: center; gap: 6px;
    background: #1E293B; border: 1px solid #334155; border-radius: 999px;
    padding: 2px 10px; font-size: 12px; margin: 0 6px 6px 0;
  }
  .sb-chip-x { color: #94A3B8; text-decoration: none; font-weight: 700; }
  .sb-chip-x:hover { color: #F87171; }

  .sb-footer { font-size: 11px; color: #64748B; text-align: center; margin-top: 10px; }

  /* Accordion */
  .accordion {
    --bs-accordion-bg: transparent;
    --bs-accordion-border-color: #1F2937;
    --bs-accordion-btn-color: #E2E8F0;
    --bs-accordion-btn-bg: #111827;
    --bs-accordion-active-bg: #0B1220;
    --bs-accordion-active-color: #93C5FD;
    --bs-accordion-btn-focus-box-shadow: 0 0 0 .2rem rgba(59,130,246,.25);
  }
  .accordion-button {
    font-size: 12px; font-weight: 600;
    letter-spacing: .06em; text-transform: uppercase;
  }
  .accordion-button::after { filter: invert(1); }

  /* Focus glow */
  .form-control:focus, .selectize-input.focus {
    box-shadow: 0 0 0 .2rem rgba(59,130,246,.25) !important;
    border-color: #3B82F6 !important;
  }

  /* Scope radio -> segmented toggle */
  .seg-toggle .shiny-options-group {
    display: flex; background: #0B0F14; border: 1px solid #1F2937;
    border-radius: 8px; padding: 3px;
  }
  .seg-toggle .form-check { flex: 1; margin: 0; padding: 0; }
  .seg-toggle .form-check-input { position: absolute; opacity: 0; }
  .seg-toggle .form-check-label {
    display: block; text-align: center; padding: 6px 0; border-radius: 6px;
    cursor: pointer; color: #94A3B8; font-size: 13px; transition: all .15s;
  }
  .seg-toggle .form-check-input:checked + .form-check-label { background: #3B82F6; color: #fff; }

  .sb-section-label {
    font-size: 11px; font-weight: 600; letter-spacing: .08em;
    text-transform: uppercase; color: #94A3B8; margin: 14px 0 6px;
  }
  .sb-section-label:first-child { margin-top: 2px; }
  .sb-hint { font-size: 11px; color: #64748B; margin-top: -4px; }
  .sb-presets .btn { flex: 1; }

  /* Range slider with typed number boxes */
  .range-wrap .irs-from, .range-wrap .irs-to, .range-wrap .irs-single {
    display: none !important;
  }
  .range-inputs { display: flex; align-items: center; gap: 8px; margin-bottom: 2px; }
  .range-inputs .form-group,
  .range-inputs .shiny-input-container {
    margin: 0 !important; flex: 1; width: auto !important;
  }
  .range-inputs input {
    text-align: center; font-weight: 600; color: #60A5FA !important;
  }
  .range-dash { color: #64748B; }
"
    ),
  
  # ==========================================================
  # SIDEBAR
  # ==========================================================
  
  sidebar = sidebar(
    width = 320,
    open = "desktop",
    
    uiOutput("sb_summary"),
    
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
          "Capacity", value = "cap", icon = icon("bolt"),
          range_slider("capacity_filter", 0, 100, c(0, 100), step = 1)
        )
      ),
      
      actionButton(
        "highlight_toggle",
        span(class = "hl-label", "Highlight coming soon"),
        class = "btn-outline-warning w-100 mt-3",
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
  
  tags$head(tags$script(HTML(APP_JS))),
  
  navset_tab(
    id = "main_tabs",
    
    # --------------------------------------------------------
    # MAP & TABLE
    # --------------------------------------------------------
    
    nav_panel(
      "Data Centers",
      
      layout_columns(
        col_widths = c(4, 4, 4),
        
        value_box(
          title = "Total capacity (MW est.)",
          value = textOutput("vb_capacity"),
          showcase = tags$span(icon("bolt"), style = "color: #FACC15;"),
          theme = "success"
        ),
        
        value_box(
          title = "Operators in view",
          value = textOutput("vb_operators"),
          showcase = tags$span(icon("building"), class = "glow-blue"),
          theme = "primary"
        ),
        
        value_box(
          title = "Cities in view",
          value = textOutput("vb_cities"),
          showcase = tags$span(icon("city"), class = "glow-blue"),
          theme = "secondary"
        )
      ),
      
      card(
        full_screen = TRUE,
        card_header("Map of Total Capacity"),
        leafletOutput("map", height = 520)
      ),
      
      card(card_header("Matching rows"), DTOutput("table"))
    ),
    
    # --------------------------------------------------------
    # UPCOMING CAPACITY
    # (filters now live in the sidebar)
    # --------------------------------------------------------
    
    nav_panel(
      "Upcoming Capacity",
      
      layout_columns(
        col_widths = c(4, 4, 4),
        
        value_box(
          title = "Total upcoming capacity (MW est.)",
          value = textOutput("vb_pipeline_mw"),
          showcase = icon("bolt"),
          theme = "success"
        ),
        
        value_box(
          title = "Upcoming entries in view",
          value = textOutput("vb_pipeline_entries"),
          showcase = tags$span(icon("clock"), class = "glow-blue"),
          theme = "warning"
        ),
        
        value_box(
          title = "Sites in view",
          value = textOutput("vb_pipeline_sites"),
          showcase = tags$span(icon("building"), class = "glow-blue"),
          theme = "primary"
        )
      ),
      
      card(
        full_screen = TRUE,
        card_header("Upcoming Capacity Map"),
        leafletOutput("pipeline_map", height = 560)
      ),
      
      card(
        card_header("Upcoming capacity by quarter"),
        plotlyOutput("pipeline_quarter_chart", height = "380px")
      ),
      
      card(
        card_header("Capacity coming available - chronological order"),
        DTOutput("pipeline_table")
      )
    ),
    
    # --------------------------------------------------------
    # VERSION HISTORY
    # --------------------------------------------------------
    
    nav_panel(
      "Version History",
      
      card(
        card_header("Refresh from DC.xlsx"),
        
        p(
          "DC.xlsx is the master file and is imported automatically ",
          "every time the app starts. Use the button below only if ",
          "you edit DC.xlsx while the app is already running and want ",
          "to reload without restarting. ",
          "The current database state is automatically preserved in version history."
        ),
        
        actionButton("refresh_excel", "Refresh from DC.xlsx", icon = icon("rotate")),
        
        br(),
        br(),
        
        uiOutput("refresh_status")
      ),
      
      card(
        card_header("History"),
        
        p(
          "Select a version below to restore it. ",
          "Restoring creates a new snapshot, so the action can be reversed."
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
  # ----------------------------------------------------------
  # Load SQLite ONLY at session start.
  # The startup import above has already run once for the whole
  # app process, so this reads fresh data.
  # ----------------------------------------------------------
  
  cur <- load_current()
  raw_data <- reactiveVal(cur)
  
  raw_pipeline <- reactiveVal(
    attach_pipeline_coordinates(load_current_pipeline(), cur)
  )
  
  # Current maximum of each range slider (used to clamp typed values)
  cap_slider_max <- reactiveVal(100)
  pipe_slider_max <- reactiveVal(100)
  
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
      
      if (length(input$state_filter) > 0) {
        keep <- keep | (df$State %in% input$state_filter)
      }
      
      if (length(input$country_filter) > 0) {
        keep <- keep | (df$Country %in% input$country_filter)
      }
      
      df <- df[keep, ]
    }
    
    df
  }
  
  # ----------------------------------------------------------
  # Update state/country/capacity filters
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
    
    df_scoped <- if (input$scope == "us") {
      df %>% filter(is_us)
    } else {
      df
    }
    
    max_cap <- suppressWarnings(max(df_scoped$Capacity_MW_est, na.rm = TRUE))
    
    if (!is.finite(max_cap)) {
      max_cap <- 100
    }
    
    cap_slider_max(ceiling(max_cap))
    
    updateSliderInput(
      session, "capacity_filter",
      max = ceiling(max_cap),
      value = c(0, ceiling(max_cap))
    )
  }, ignoreNULL = FALSE)
  
  # ----------------------------------------------------------
  # Upcoming Capacity country choices
  # Based ONLY on raw_pipeline(); independent of the map scope.
  # ----------------------------------------------------------
  
  observeEvent(raw_pipeline(), {
    countries <- raw_pipeline() %>%
      filter(!is.na(Country), Country != "") %>%
      distinct(Country) %>%
      arrange(Country) %>%
      pull(Country)
    
    updateSelectizeInput(
      session, "pipeline_country_filter",
      choices = countries,
      server = TRUE
    )
    
    pl_states <- raw_pipeline() %>%
      filter(is_us, !is.na(State), State != "") %>%
      distinct(State) %>%
      arrange(State) %>%
      pull(State)
    
    updateSelectizeInput(
      session, "pipeline_state_filter",
      choices = pl_states,
      server = TRUE
    )
  }, ignoreNULL = FALSE)
  
  # ----------------------------------------------------------
  # City choices
  # ----------------------------------------------------------
  
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
  
  # ----------------------------------------------------------
  # Operator choices
  # ----------------------------------------------------------
  
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
  
  fmt <- function(x) format(round(x), big.mark = ",")
  
  # Live results summary (changes with the active tab)
  output$sb_summary <- renderUI({
    if (identical(input$main_tabs, "Upcoming Capacity")) {
      df <- filtered_pipeline()
      n <- nrow(df)
      total <- nrow(raw_pipeline())
      mw <- sum(df$MW_available, na.rm = TRUE)
      label <- "upcoming entries"
    } else {
      df <- filtered()
      n <- nrow(df)
      total <- nrow(raw_data())
      mw <- sum(df$Capacity_MW_est, na.rm = TRUE)
      label <- "sites"
    }
    
    div(
      class = "sb-summary",
      div(class = "sb-summary-main",
          tags$b(fmt(n)), paste0(" of ", fmt(total), " ", label)),
    )
  })
  
  # Removable filter chips
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
  
  # Presets
  # Highlight toggle (button at the bottom of the sidebar)
  highlight_on <- reactiveVal(FALSE)
  
  observeEvent(input$highlight_toggle, {
    highlight_on(!highlight_on())
  })
  
  observeEvent(highlight_on(), {
    session$sendCustomMessage("hl_state", highlight_on())
  }, ignoreInit = TRUE)
  
  # Slider <-> typed-number boxes (drag OR type)
  sync_range <- function(id, max_fn) {
    lo_id <- paste0(id, "_lo")
    hi_id <- paste0(id, "_hi")
    
    # slider -> boxes
    observeEvent(input[[id]], {
      updateNumericInput(session, lo_id, value = input[[id]][1])
      updateNumericInput(session, hi_id, value = input[[id]][2])
    })
    
    # boxes -> slider
    observeEvent(list(input[[lo_id]], input[[hi_id]]), {
      lo <- input[[lo_id]]
      hi <- input[[hi_id]]
      req(is.numeric(lo), is.numeric(hi), is.finite(lo), is.finite(hi))
      
      mx <- max_fn()
      a <- min(max(min(lo, hi), 0), mx)
      b <- min(max(max(lo, hi), 0), mx)
      
      # Show the clamped / re-ordered values back in the boxes
      if (a != lo || b != hi) {
        updateNumericInput(session, lo_id, value = a)
        updateNumericInput(session, hi_id, value = b)
      }
      
      cur <- input[[id]]
      if (is.null(cur) || abs(cur[1] - a) > 1e-9 || abs(cur[2] - b) > 1e-9) {
        updateSliderInput(session, id, value = c(a, b))
      }
    }, ignoreInit = TRUE)
  }
  
  sync_range("capacity_filter", cap_slider_max)
  sync_range("pipeline_mw_filter", pipe_slider_max)
  
  # Data freshness footer
  output$data_freshness <- renderText({
    refresh_trigger()
    
    v <- list_versions() %>%
      filter(!grepl("^Auto-snapshot", note))
    
    if (nrow(v) == 0) return("No data loaded yet")
    
    paste0("Data refreshed ", substr(v$timestamp[1], 1, 16), " \u00b7 v", v$version[1])
  })
  
  # ----------------------------------------------------------
  # Reset (defaults to Global, always resets capacity sliders)
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
    
    gmax <- suppressWarnings(max(raw_data()$Capacity_MW_est, na.rm = TRUE))
    if (!is.finite(gmax)) gmax <- 100
    
    updateSliderInput(
      session, "capacity_filter",
      max = ceiling(gmax),
      value = c(0, ceiling(gmax))
    )
    
    pmax_mw <- suppressWarnings(max(raw_pipeline()$MW_available, na.rm = TRUE))
    if (!is.finite(pmax_mw)) pmax_mw <- 100
    
    updateSliderInput(
      session, "pipeline_mw_filter",
      max = ceiling(pmax_mw),
      value = c(0, ceiling(pmax_mw))
    )
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
    
    df %>%
      filter(
        is.na(Capacity_MW_est) |
          (Capacity_MW_est >= input$capacity_filter[1] &
             Capacity_MW_est <= input$capacity_filter[2])
      )
  })
  
  # ----------------------------------------------------------
  # Pipeline filtering
  #
  # Intentionally independent of the Data Centers filters:
  # - US-only scope does NOT remove Indonesia
  # - State / country / city / operator filters do NOT affect it
  # It has its own country, quarter, and MW filters.
  # ----------------------------------------------------------
  
  filtered_pipeline <- reactive({
    df <- raw_pipeline()
    
    # Upcoming Capacity has its own US / Global scope.
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
    
    df %>%
      arrange(Quarter_Date, Operator, Country, State, City_clean)
  })
  
  # ----------------------------------------------------------
  # Set maximum for upcoming MW slider
  # ----------------------------------------------------------
  
  observeEvent(raw_pipeline(), {
    max_mw <- suppressWarnings(max(raw_pipeline()$MW_available, na.rm = TRUE))
    
    if (!is.finite(max_mw)) {
      max_mw <- 100
    }
    
    pipe_slider_max(ceiling(max_mw))
    
    updateSliderInput(
      session, "pipeline_mw_filter",
      max = ceiling(max_mw),
      value = c(0, ceiling(max_mw))
    )
    
  }, ignoreNULL = FALSE)
  
  # ----------------------------------------------------------
  # Year shortcut buttons for the Quarter filter
  # ----------------------------------------------------------
  
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
  
  # ==========================================================
  # UPCOMING HIGHLIGHT LOGIC
  # ==========================================================
  
  upcoming_soon_keys <- reactive({
    pl <- raw_pipeline()
    
    if (nrow(pl) == 0) {
      return(character(0))
    }
    
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
  
  # ==========================================================
  # VALUE BOXES
  # ==========================================================
  
  output$vb_operators <- renderText({
    n_distinct(filtered()$Operator)
  })
  
  output$vb_cities <- renderText({
    n_distinct(filtered()$City_clean)
  })
  
  output$vb_capacity <- renderText({
    total <- sum(filtered()$Capacity_MW_est, na.rm = TRUE)
    format(round(total), big.mark = ",")
  })
  
  output$vb_pipeline_entries <- renderText({
    nrow(filtered_pipeline())
  })
  
  output$vb_pipeline_sites <- renderText({
    filtered_pipeline() %>%
      distinct(Operator, City_clean, State, Country) %>%
      nrow()
  })
  
  output$vb_pipeline_mw <- renderText({
    format(round(sum(filtered_pipeline()$MW_available, na.rm = TRUE)), big.mark = ",")
  })
  
  # ==========================================================
  # MAP
  # ==========================================================
  
  output$map <- renderLeaflet({
    leaflet(options = leafletOptions(
      worldCopyJump = FALSE,
      minZoom = 2,
      maxZoom = 18
    )) %>%
      addTiles(
        urlTemplate = CARTO_POSITRON_URL,
        attribution = CARTO_ATTRIBUTION,
        options = tileOptions(noWrap = TRUE)
      ) %>%
      setMaxBounds(lng1 = -180, lat1 = -85, lng2 = 180, lat2 = 85) %>%
      setView(lng = -98.5, lat = 39.5, zoom = 4) %>%
      addEasyButton(reset_view_button("map_reset"))
  })
  
  # Recenters the map when scope changes (US or global)
  # Recenters the map (used by the scope toggle and the reset button)
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
  ##############################
  observe({
    df <- filtered() %>%
      filter(!is.na(Latitude), !is.na(Longitude))
    
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
        
        op_txt = htmltools::htmlEscape(enc2utf8(coalesce(Operator, "Unknown operator"))),
        city_txt = htmltools::htmlEscape(enc2utf8(coalesce(City_clean, ""))),
        place_txt = htmltools::htmlEscape(enc2utf8(
          if_else(is.na(State) | State == "", coalesce(Country, ""), State)
        )),
        cap_txt = htmltools::htmlEscape(enc2utf8(coalesce(Capacity, "n/a"))),
        cool_txt = htmltools::htmlEscape(enc2utf8(coalesce(Cooling, "Unknown"))),
        
        popup_html = paste0(
          "<div style='min-width:200px'>",
          "<b>", op_txt, "</b><br>",
          city_txt, ", ", place_txt, "<br>",
          "Capacity: ", cap_txt,
          if_else(is.na(Cooling), "", paste0("<br>Cooling: ", cool_txt)),
          if_else(
            is_upcoming_soon,
            "<br><b style='color:#F59E0B'>Capacity coming available soon - see Upcoming Capacity tab</b>",
            ""
          ),
          "</div>"
        ),
        label_txt = paste0(op_txt, " \u2014 ", city_txt)
      )
    
    leafletProxy("map", data = df) %>%
      clearMarkers() %>%
      clearMarkerClusters() %>%
      addCircleMarkers(
        lng = ~map_lng,
        lat = ~map_lat,
        radius = 8,
        fillOpacity = 0.85,
        color = ~ifelse(is_upcoming_soon, "#F59E0B", "#60A5FA"),
        fillColor = ~ifelse(is_upcoming_soon, "#F59E0B", "#3B82F6"),
        weight = 2,
        popup = ~popup_html,
        label = ~lapply(label_txt, htmltools::HTML),
        clusterOptions = markerClusterOptions(
          disableClusteringAtZoom = 14,
          spiderfyOnMaxZoom = FALSE,
          maxClusterRadius = 40
        )
      )
  })
  
  # ==========================================================
  # UPCOMING CAPACITY MAP
  # ==========================================================
  
  output$pipeline_map <- renderLeaflet({
    # Rendered from the SAME reactive dataset as the Upcoming Capacity table.
    df <- filtered_pipeline()
    
    base_map <- leaflet(options = leafletOptions(
      worldCopyJump = FALSE,
      minZoom = 2,
      maxZoom = 18
    )) %>%
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
    
    if (nrow(df) == 0) {
      return(base_map)
    }
    
    # Give every upcoming row a location. Prefer exact geocoded
    # coordinates; otherwise use the country centroid.
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
    
    # If several upcoming rows are at the same location, spread
    # them slightly so each quarter/capacity row can be clicked.
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
    
    if (nrow(map_df) == 0) {
      return(base_map)
    }
    
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
    
    popup_text <- paste0(
      "<div style='min-width:210px'>",
      "<b style='font-size:15px'>",
      ifelse(is.na(map_df$Operator), "Unknown operator", map_df$Operator),
      "</b><br>",
      ifelse(
        is.na(map_df$City_clean) | map_df$City_clean == "",
        map_df$Country,
        map_df$City_clean
      ),
      ifelse(
        !is.na(map_df$State) & map_df$State != "",
        paste0(", ", map_df$State),
        ""
      ),
      "<br><br>",
      "<b>Coming online:</b> ", map_df$Quarter, "<br>",
      "<b>MW available:</b> ",
      format(round(map_df$MW_available, 1), big.mark = ","), " MW",
      "</div>"
    )
    
    base_map %>%
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
  })
  
  # ==========================================================
  # CURRENT TABLE
  # ==========================================================
  
  output$table <- renderDT({
    filtered() %>%
      select(
        Operator,
        City = City_clean,
        State,
        Country,
        Region,
        Capacity,
        `Capacity (MW est.)` = Capacity_MW_est
      ) %>%
      arrange(Operator, Country, State, City)
  }, options = list(pageLength = 15), rownames = FALSE)
  
  ##
  # UPCOMING CAPACITY BY QUARTER
  # ==========================================================
  
  output$pipeline_quarter_chart <- renderPlotly({
    
    df <- filtered_pipeline()
    
    if (nrow(df) == 0) {
      return(
        plot_ly() %>%
          layout(
            paper_bgcolor = "#111827",
            plot_bgcolor = "#111827",
            font = list(color = "#E2E8F0", family = "Inter"),
            xaxis = list(visible = FALSE),
            yaxis = list(visible = FALSE),
            annotations = list(
              list(
                text = "No upcoming capacity matches the current filters.",
                x = 0.5,
                y = 0.5,
                xref = "paper",
                yref = "paper",
                showarrow = FALSE,
                font = list(size = 14, color = "#94A3B8")
              )
            )
          )
      )
    }
    
    # Total MW by quarter
    quarter_totals <- df %>%
      mutate(
        Quarter = factor(Quarter, levels = QUARTER_COLS),
        Operator = if_else(
          is.na(Operator) | Operator == "",
          "Unknown operator",
          Operator
        )
      ) %>%
      group_by(Quarter) %>%
      summarise(
        MW = sum(MW_available, na.rm = TRUE),
        .groups = "drop"
      )
    
    # Operator breakdown for hover
    operator_breakdown <- df %>%
      mutate(
        Quarter = factor(Quarter, levels = QUARTER_COLS),
        Operator = if_else(
          is.na(Operator) | Operator == "",
          "Unknown operator",
          Operator
        )
      ) %>%
      group_by(Quarter, Operator) %>%
      summarise(
        MW = sum(MW_available, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      group_by(Quarter) %>%
      summarise(
        Breakdown = paste0(
          "<b>", Operator, "</b>: ",
          format(round(MW, 1), big.mark = ","),
          " MW",
          collapse = "<br>"
        ),
        .groups = "drop"
      )
    
    chart_data <- quarter_totals %>%
      left_join(operator_breakdown, by = "Quarter")
    
    plot_ly(
      data = chart_data,
      x = ~Quarter,
      y = ~MW,
      type = "bar",
      
      text = ~format(round(MW, 1), big.mark = ","),
      textposition = "outside",
      
      hovertext = ~paste0(
        "<b>", Quarter, "</b>",
        "<br><b>Total: ", format(round(MW, 1), big.mark = ","), " MW</b>",
        "<br><br>",
        Breakdown
      ),
      
      hoverinfo = "text",
      
      marker = list(
        color = "#3B82F6"
      )
    ) %>%
      
      layout(
        paper_bgcolor = "#111827",
        plot_bgcolor = "#111827",
        
        font = list(
          color = "#E2E8F0",
          family = "Inter"
        ),
        
        title = list(
          text = "Capacity coming online by quarter",
          font = list(
            size = 16,
            color = "#F9FAFB"
          )
        ),
        
        xaxis = list(
          title = NULL,
          categoryorder = "array",
          categoryarray = QUARTER_COLS,
          tickfont = list(color = "#CBD5E1"),
          gridcolor = "#1F2937",
          linecolor = "#374151"
        ),
        
        yaxis = list(
          title = "Upcoming capacity (MW)",
          titlefont = list(color = "#CBD5E1"),
          tickfont = list(color = "#CBD5E1"),
          gridcolor = "#1F2937",
          zerolinecolor = "#374151"
        ),
        
        hovermode = "closest",
        
        margin = list(
          l = 65,
          r = 30,
          t = 65,
          b = 60
        ),
        
        showlegend = FALSE
      ) %>%
      
      config(
        displayModeBar = FALSE,
        responsive = TRUE
      )
  })
  
  # ==========================================================
  # UPCOMING CAPACITY TABLE
  # ==========================================================
  
  output$pipeline_table <- renderDT({
    df <- filtered_pipeline()
    
    if (nrow(df) == 0) {
      return(datatable(
        tibble(Message = "No upcoming capacity matches the current filters."),
        rownames = FALSE,
        options = list(dom = "t")
      ))
    }
    
    # Long format: each Excel quarter is its own row, so every
    # quarter from Q3 2026 onward appears as a chronological entry.
    df <- df %>%
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
    
    datatable(
      df,
      rownames = FALSE,
      options = list(
        pageLength = 25,
        scrollX = TRUE,
        autoWidth = TRUE,
        order = list(list(7, "asc")),
        columnDefs = list(list(visible = FALSE, targets = 7))
      )
    )
  })
  
  # ==========================================================
  # REFRESH FROM DC.XLSX
  # ==========================================================
  
  refresh_trigger <- reactiveVal(0)
  
  output$refresh_status <- renderUI({
    refresh_trigger()
    
    versions <- list_versions()
    
    if (nrow(versions) == 0) {
      return(tags$p("No Excel refresh has been performed yet."))
    }
    
    latest <- versions %>%
      arrange(desc(version)) %>%
      slice(1)
    
    tags$p(
      strong("Latest database update: "),
      latest$timestamp,
      tags$br(),
      "Version: ", latest$version,
      tags$br(),
      "Sites: ", latest$n_rows,
      tags$br(),
      "Upcoming entries: ", latest$n_pipeline_rows
    )
  })
  
  # Step 1: the button only opens a password prompt.
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
  
  # Step 2: the refresh only runs if the password matches.
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
        type = "error",
        duration = 10
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
      
      # Update reactive data immediately.
      raw_data(load_current())
      raw_pipeline(attach_pipeline_coordinates(load_current_pipeline(), load_current()))
      refresh_trigger(refresh_trigger() + 1)
      
      showNotification(
        paste0(
          "DC.xlsx imported successfully: ",
          nrow(result$current), " sites and ",
          nrow(result$pipeline), " upcoming-capacity entries."
        ),
        type = "message",
        duration = 8
      )
    }, error = function(e) {
      showNotification(
        paste("Excel refresh failed:", conditionMessage(e)),
        type = "error",
        duration = 12
      )
    })
  })
  # ==========================================================
  # VERSION HISTORY TABLE
  # ==========================================================
  
  output$version_table <- renderDT({
    refresh_trigger()
    
    versions <- list_versions()
    
    if (nrow(versions) == 0) {
      return(datatable(tibble(Message = "No versions yet.")))
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
        options = list(pageLength = 10)
      )
  })
  
  # ==========================================================
  # RESTORE VERSION
  # ==========================================================
  
  observeEvent(input$version_table_rows_selected, {
    sel <- input$version_table_rows_selected
    
    req(sel)
    
    versions <- list_versions() %>%
      arrange(desc(version))
    
    v <- versions$version[sel]
    
    showModal(modalDialog(
      title = paste("Restore version", v, "?"),
      
      p(paste0(
        "This will make version ", v, " live again. ",
        "The current data will be snapshotted first, so this is safe to undo."
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
    # Wrong password: keep the dialog open so they can retry.
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