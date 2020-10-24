#
# Merge Microsoft Powerpoint Presentations into one Presenation
#
# Source the whole script first to load the functions defined below
#
# Install required package:
# remotes::install_github("KWB-R/kwb.utils")

# MAIN -------------------------------------------------------------------------
if (FALSE)
{
  ppt_file <- merge_ppt_files_below_path(
    path = "~/../Downloads/P/presentations_sema-berlin-2",
    pattern = "^2019_04.*\\.pptx$"
  )
  
  kwb.utils::hsOpenWindowsExplorer(dirname(ppt_file))
}

# merge_ppt_files_below_path ---------------------------------------------------
merge_ppt_files_below_path <- function(path, pattern = "\\.pptx?$")
{
  files <- list_ppt_recursively(path, pattern)
  
  if (length(files) == 0L) {
    return()
  }
  
  merge_ppt_files(files)
}

# list_ppt_recursively ---------------------------------------------------------
list_ppt_recursively <- function(
  path, pattern = "\\.pptx?$", ignore.case = TRUE
)
{
  files <- dir(
    path, 
    pattern = pattern, 
    full.names = TRUE,
    recursive = TRUE, 
    ignore.case = ignore.case
  )
  
  if (length(files) == 0L) {
    
    message(sprintf(
      "No files matching '%s' found below '%s'.", pattern, path
    ))
  }
  
  files  
}

# merge_ppt_files --------------------------------------------------------------
merge_ppt_files <- function(files)
{
  paths <- kwb.utils::resolve(list(
    input = "<app>/input",
    app = "<temp>/ppt-merge",
    script = "<app>/merge.vbs",
    temp = tempdir()
  ))
  
  delete_folder_recursively(paths$app, really = TRUE, force = TRUE)
  
  idir <- kwb.utils::createDirectory(paths$input, dbg = FALSE)
  
  copy_files_to_dir(files, to = paths$input)
  
  script_file <- "merged.ppt"
  
  kwb.utils::catAndRun(sprintf("Writing VB script to '%s'", paths$script), {
    writeLines(con = paths$script, kwb.utils::resolve(
      paste(readLines("script-template.txt"), collapse = "\n"),
      PPT_MERGE_FILE = script_file,
      PPT_MERGE_FOLDER = paths$input
    ))
  })
  
  system(paste("WScript", paths$script), wait = TRUE)
  
  file.path(paths$app, script_file)
}

# delete_folder_recursively ----------------------------------------------------
delete_folder_recursively <- function(path, really = FALSE, force = FALSE)
{
  stopifnot(is.character(path), length(path) == 1L)
  
  if (! dir.exists(path)) {
    message("Folder '", path, "' does not exist. Nothing to delete.")
    return()
  }
  
  if (! really) {
    message("Sorry, I do not delete '", path, "' unless you really want it.")
    return()
  }
    
  kwb.utils::catAndRun(sprintf("Deleting '%s' recursively", path), {
    unlink(path, recursive = TRUE, force = force)
  })
  
  if (dir.exists(path)) {
    stop("The folder '", path, "' could not be deleted.", call. = FALSE)
  }
}

# copy_files_to_dir ------------------------------------------------------------
copy_files_to_dir <- function(from, to)
{
  stopifnot(is.character(to), length(to) == 1L, dir.exists(to))

  kwb.utils::catAndRun(
    sprintf("Copying %d files to '%s'", length(from), to), 
    success <- file.copy(from, to)
  )
  
  if (any(! success)) {
    
    stop(
      sprintf("Could not copy the following files to '%s':\n* ", to),
      paste(from[! success], collapse = "\n* "),
      call. = FALSE
    )
  }
  
  file.path(to, basename(from))
}
