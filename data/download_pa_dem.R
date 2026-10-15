# Download USGS 3DEP 1/3 arc-second (about 10 m) DEM tiles covering Pennsylvania.
# Files are saved under data/dem_pa_1_3_arc_second/.

# Install these packages once if needed:
# install.packages(c("httr2", "jsonlite"))

library(httr2)
library(jsonlite)

# Pennsylvania's approximate bounding box in longitude/latitude (WGS84).
# The API returns whole 1 x 1 degree tiles that overlap this box, so a few
# neighboring-state tiles may be included along the edges.
pa_bbox <- "-80.52,39.72,-74.69,42.27"

api_url <- "https://tnmaccess.nationalmap.gov/api/v1/products"
dem_dataset <- "National Elevation Dataset (NED) 1/3 arc-second"

response <- request(api_url) |>
  req_url_query(
    datasets = dem_dataset,
    bbox = pa_bbox,
    prodFormats = "GeoTIFF",
    max = 1000
  ) |>
  req_timeout(120) |>
  req_perform()

products <- fromJSON(resp_body_string(response), simplifyDataFrame = TRUE)

if (length(products$items) == 0) {
  stop("The USGS API returned no DEM tiles. Check the API response and try again.")
}

tiles <- products$items

# TNM may list historical versions of a tile. Keep only the newest publication
# date for each 1 x 1 degree tile to avoid downloading duplicate versions.
tile_names <- regmatches(
  tiles$downloadURL,
  regexpr("n[0-9]{2}w[0-9]{3}", tiles$downloadURL)
)
if (any(tile_names == "")) {
  stop("Could not identify a tile name in one or more USGS download URLs.")
}

tiles$publicationDate <- as.Date(tiles$publicationDate)
tiles <- tiles[order(tiles$publicationDate, decreasing = TRUE), ]
tiles <- tiles[!duplicated(tile_names[order(tiles$publicationDate, decreasing = TRUE)]), ]

output_dir <- file.path("data", "dem_pa_1_3_arc_second")
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

for (i in seq_len(nrow(tiles))) {
  download_url <- tiles$downloadURL[[i]]
  filename <- basename(sub("[?#].*$", "", download_url))
  destination <- file.path(output_dir, filename)

  if (file.exists(destination) && file.info(destination)$size > 0) {
    message("Already downloaded: ", filename)
    next
  }

  message(sprintf("Downloading %d of %d: %s", i, nrow(tiles), filename))
  tryCatch(
    download.file(download_url, destination, mode = "wb", method = "libcurl", quiet = FALSE),
    error = function(e) {
      if (file.exists(destination)) unlink(destination)
      stop("Download failed for ", filename, ": ", conditionMessage(e))
    }
  )
}

message("Finished. DEM tiles are in: ", normalizePath(output_dir))
