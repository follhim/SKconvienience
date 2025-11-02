#' @importFrom utils packageVersion
.onAttach <- function(libname, pkgname) {
  packageStartupMessage(
    "To cite SKconvienience in publications, use:\n\n",
    "  Kim, S. (2025). SKconvienience: A package for convenience in statistics\n",
    "  (Version ", packageVersion("SKconvienience"), "). ",
    "https://follhim.github.io/SKconvienience/\n\n",
    "A BibTeX entry for LaTeX users is available with citation('SKconvienience')"
  )
}
