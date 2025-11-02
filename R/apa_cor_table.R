#' APA 7 Style Correlation Table with Shifted Descriptives (Lower Triangle) - Excel Export
#'
#' Creates a properly formatted APA 7 style correlation table with superscript asterisks
#' and proper lower triangle orientation when descriptives are shifted, exported to Excel
#'
#' @param data Data frame containing variables for correlation
#' @param filename Output filename (should end in .xlsx)
#' @param show.conf.interval Whether to display confidence intervals
#' @param show.sig.stars Whether to display significance stars
#' @param show.pvalue Whether to display p-values
#' @param landscape Whether to create landscape-oriented output
#' @param shift.descriptives Whether to display descriptives as rows at bottom
#' @param cor.methods Named vector specifying correlation methods for variables
#' @return An object with the correlation table data
#' @export
#' @examples
#' # Basic usage with sample data
#' test_data <- data.frame(
#'   x = rnorm(30),
#'   y = rnorm(30),
#'   z = rnorm(30)
#' )
#' apa_cor_table(test_data)
#'
#' # With shifted descriptives
#' apa_cor_table(test_data, shift.descriptives = TRUE)
#' @importFrom stats cor.test sd
apa_cor_table <- function(data,
                          filename = NA,
                          show.conf.interval = TRUE,
                          show.sig.stars = TRUE,
                          show.pvalue = TRUE,
                          landscape = TRUE,
                          shift.descriptives = FALSE,
                          cor.methods = NULL) {

  # Check if required packages are installed
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    if (!is.na(filename)) {
      warning("Package 'openxlsx' is required for Excel output.\n",
              "Install it using: install.packages('openxlsx')\n",
              "Returning table object without saving file.")
    }
  }

  # HELPER FUNCTIONS

  # Format numbers according to APA style
  txt_number <- function(x, digits = 2) {
    if (is.na(x)) {
      return(NA)
    }

    if (abs(x) < 0.001) {
      return("< .001")
    }

    if (abs(x) < 1) {
      formatted <- format(round(x, digits), nsmall = digits)
      return(sub("0\\.", "\\.", formatted))
    } else {
      return(format(round(x, digits), nsmall = digits))
    }
  }

  # Format confidence intervals
  txt_ci <- function(ctest) {
    ci <- ctest$conf.int
    ci.text <- paste("[", txt_number(ci[1]), ", ", txt_number(ci[2]), "]", sep = "")
    return(ci.text)
  }

  # Format correlation coefficient with stars
  txt_r <- function(ctest, show_stars = TRUE) {
    if (show_stars) {
      if (ctest$p.value < 0.001) {
        return(paste(txt_number(ctest$estimate), "***", sep = ""))
      } else if (ctest$p.value < 0.01) {
        return(paste(txt_number(ctest$estimate), "**", sep = ""))
      } else if (ctest$p.value < 0.05) {
        return(paste(txt_number(ctest$estimate), "*", sep = ""))
      } else {
        return(txt_number(ctest$estimate))
      }
    } else {
      return(txt_number(ctest$estimate))
    }
  }

  # Format p-values for display
  txt_p <- function(ctest) {
    if (ctest$p.value < 0.001) {
      return("p < .001")
    } else if (ctest$p.value < 0.01) {
      return("p < .01")
    } else if (ctest$p.value < 0.05) {
      return("p < .05")
    } else {
      return(paste("p = ", txt_number(ctest$p.value), sep = ""))
    }
  }

  # Calculate correlation based on method
  calculate_correlation <- function(x, y, method = "pearson") {
    if (method == "spearman") {
      return(cor.test(x, y, method = "spearman"))
    } else if (method == "pointbiserial") {
      # Point-biserial is technically Pearson with a dichotomous variable
      return(cor.test(x, y))
    } else {
      # Default to Pearson
      return(cor.test(x, y))
    }
  }

  # Process input data
  data <- as.data.frame(data)

  # Keep only numeric columns
  df_col <- dim(data)[2]
  column_is_numeric <- c()
  for (i in 1:df_col) {
    column_is_numeric[i] <- is.numeric(data[,i])
  }

  if (!all(column_is_numeric)) {
    warning("Non-numeric columns have been removed from the analysis.")
    data <- data[,column_is_numeric]
  }

  # Get dimensions
  number_variables <- ncol(data)
  number_columns <- number_variables - 1

  # Initialize matrices for output
  output_cor <- matrix(" ", number_variables, number_variables)
  output_ci <- matrix(" ", number_variables, number_variables)
  output_pvalue <- matrix(" ", number_variables, number_variables)
  output_descriptives <- matrix(" ", number_variables, 3)

  # Create variable names
  var_names <- names(data)
  output_variable_names <- paste(as.character(1:number_variables), ". ", var_names, sep="")

  # Calculate descriptive statistics
  for (i in 1:number_variables) {
    output_descriptives[i,1] <- as.character(sum(!is.na(data[,i]), na.rm=TRUE))
    output_descriptives[i,2] <- txt_number(mean(data[,i], na.rm=TRUE))
    output_descriptives[i,3] <- txt_number(sd(data[,i], na.rm=TRUE))
  }

  # Calculate correlations
  for (i in 1:number_variables) {
    for (j in 1:number_variables) {
      if (j < i) {
        x <- data[,i]
        y <- data[,j]

        # Determine correlation method
        method <- "pearson" # default
        if (!is.null(cor.methods)) {
          if (var_names[i] %in% names(cor.methods)) {
            method <- cor.methods[var_names[i]]
          }
          if (var_names[j] %in% names(cor.methods)) {
            method <- cor.methods[var_names[j]]
          }
        }

        # Calculate correlation
        ctest <- tryCatch({
          calculate_correlation(x, y, method)
        }, error = function(e) {
          warning(paste("Error calculating correlation between", var_names[i], "and", var_names[j], ":", e$message))
          list(estimate = NA, p.value = NA, conf.int = c(NA, NA))
        })

        # Format outputs
        output_cor[i,j] <- txt_r(ctest, show.sig.stars)
        output_ci[i,j] <- txt_ci(ctest)
        output_pvalue[i,j] <- txt_p(ctest)
      }
    }
  }

  # Create output table
  if (!shift.descriptives) {
    # Standard format with descriptives in columns

    # Create a data frame directly
    df_out <- data.frame(Variable = character(number_variables),
                         N = character(number_variables),
                         M = character(number_variables),
                         SD = character(number_variables),
                         stringsAsFactors = FALSE)

    # Add correlation columns
    for (j in 1:number_variables) {
      df_out[paste0(j)] <- ""
    }

    # Fill in the data frame
    for (i in 1:number_variables) {
      # Variable name and descriptives
      df_out$Variable[i] <- output_variable_names[i]
      df_out$N[i] <- output_descriptives[i,1]
      df_out$M[i] <- output_descriptives[i,2]
      df_out$SD[i] <- output_descriptives[i,3]

      # Correlations
      for (j in 1:number_variables) {
        if (j < i) {
          df_out[i, j+4] <- output_cor[i,j]
        }
      }
    }

    # Format column names properly
    colnames(df_out) <- c("Variable", "N", "M", "SD", as.character(1:number_variables))

    # Store as output matrix for consistency
    output_matrix <- as.matrix(df_out)
  }
  else {
    # IMPROVED shifted descriptives with LOWER triangle (not upper)

    # Create a data frame for the correlation matrix
    df_out <- data.frame(Variable = character(number_variables),
                         stringsAsFactors = FALSE)

    # Add correlation columns
    for (j in 1:number_variables) {
      df_out[paste0(j)] <- ""
    }

    # Fill in the data frame with variable names
    for (i in 1:number_variables) {
      df_out$Variable[i] <- output_variable_names[i]
    }

    # Fill in LOWER triangle correlations
    for (i in 1:number_variables) {
      for (j in 1:number_variables) {
        if (j < i) {  # This is for lower triangle
          df_out[i, j+1] <- output_cor[i,j]
        }
      }
    }

    # Add descriptive statistics rows at the bottom
    n_row <- c("N", output_descriptives[,1])
    m_row <- c("M", output_descriptives[,2])
    sd_row <- c("SD", output_descriptives[,3])

    # Combine into final data frame
    df_out_with_desc <- rbind(
      df_out,
      n_row,
      m_row,
      sd_row
    )

    # Format column names properly
    colnames(df_out_with_desc) <- c("Variable", as.character(1:number_variables))

    # Store as output matrix for consistency
    output_matrix <- as.matrix(df_out_with_desc)
  }

  # Create table title and notes
  table_title <- "Descriptive Statistics and Correlations"

  # Create APA style notes
  note_text <- "Note. N = number of cases. M = mean. SD = standard deviation."

  if (show.sig.stars) {
    sig_note <- "* indicates p < .05. ** indicates p < .01. *** indicates p < .001."
  }

  # Add note about correlation methods if provided
  if (!is.null(cor.methods) && length(cor.methods) > 0) {
    methods_note <- "Correlation methods: "
    for (i in 1:length(cor.methods)) {
      var_name <- names(cor.methods)[i]
      method <- cor.methods[i]
      methods_note <- paste(methods_note, var_name, " (", method, ")",
                            ifelse(i < length(cor.methods), ", ", ""), sep="")
    }
  }

  # Create return object
  result <- list(
    table.title = table_title,
    table.body = output_matrix,
    table.note = note_text,
    landscape = landscape
  )

  class(result) <- "apa.cor.table"

  # Print table to console
  cat("\n", table_title, "\n\n")
  print(output_matrix, quote = FALSE)
  cat("\n", note_text, "\n")
  if (show.sig.stars) cat(sig_note, "\n")
  if (!is.null(cor.methods) && length(cor.methods) > 0) cat(methods_note, "\n")

  # Create and save Excel file if filename is provided
  if (!is.na(filename)) {
    if (!requireNamespace("openxlsx", quietly = TRUE)) {
      warning("Package 'openxlsx' is required for Excel output.")
    } else {
      # Ensure filename ends with .xlsx
      if (grepl("\\.xls$", filename) && !grepl("\\.xlsx$", filename)) {
        filename <- sub("\\.xls$", ".xlsx", filename)
        warning("Changed file extension from .xls to .xlsx for better compatibility.")
      } else if (!grepl("\\.xlsx$", filename)) {
        filename <- paste0(filename, ".xlsx")
      }

      # Create workbook
      wb <- openxlsx::createWorkbook()

      # Add worksheet
      sheet_name <- "Correlation Table"
      openxlsx::addWorksheet(wb, sheet_name, orientation = ifelse(landscape, "landscape", "portrait"))

      # Convert output matrix to data frame for Excel export
      df_table <- as.data.frame(output_matrix, stringsAsFactors = FALSE)

      # Add title
      openxlsx::writeData(wb, sheet_name, table_title, startRow = 1, startCol = 1)

      # Add main table (starting from row 3 to leave space for title)
      openxlsx::writeData(wb, sheet_name, df_table, startRow = 3, startCol = 1,
                          colNames = TRUE, rowNames = FALSE)

      # Calculate positions for notes
      table_end_row <- 3 + nrow(df_table) + 1
      note_start_row <- table_end_row + 1

      # Add notes
      openxlsx::writeData(wb, sheet_name, note_text, startRow = note_start_row, startCol = 1)

      if (show.sig.stars) {
        openxlsx::writeData(wb, sheet_name, sig_note, startRow = note_start_row + 1, startCol = 1)
      }

      if (!is.null(cor.methods) && length(cor.methods) > 0) {
        methods_row <- ifelse(show.sig.stars, note_start_row + 2, note_start_row + 1)
        openxlsx::writeData(wb, sheet_name, methods_note, startRow = methods_row, startCol = 1)
      }

      # Style the table according to APA 7 guidelines

      # Set font to Times New Roman for entire sheet
      openxlsx::modifyBaseFont(wb, fontSize = 12, fontName = "Times New Roman")

      # Style title (bold and centered)
      title_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        textDecoration = "bold",
        halign = "center"
      )
      openxlsx::addStyle(wb, sheet_name, title_style, rows = 1, cols = 1:ncol(df_table))

      # Style headers
      header_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        halign = "center",
        border = "top-bottom",
        borderStyle = "thin"
      )
      openxlsx::addStyle(wb, sheet_name, header_style, rows = 3, cols = 1:ncol(df_table))

      # Style data cells
      data_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        halign = "center"
      )

      # Apply center alignment to all data cells except first column
      if (ncol(df_table) > 1) {
        openxlsx::addStyle(wb, sheet_name, data_style,
                           rows = 4:(3 + nrow(df_table)),
                           cols = 2:ncol(df_table),
                           gridExpand = TRUE)
      }

      # Left align first column (variable names)
      left_align_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        halign = "left"
      )
      openxlsx::addStyle(wb, sheet_name, left_align_style,
                         rows = 4:(3 + nrow(df_table)),
                         cols = 1,
                         gridExpand = TRUE)

      # Add bottom border to last row of data
      bottom_border_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        border = "bottom",
        borderStyle = "thin"
      )
      openxlsx::addStyle(wb, sheet_name, bottom_border_style,
                         rows = 3 + nrow(df_table),
                         cols = 1:ncol(df_table))

      # Style notes
      note_style <- openxlsx::createStyle(
        fontSize = 12,
        fontName = "Times New Roman",
        halign = "left"
      )

      # Calculate how many note rows we actually have
      note_rows <- c(note_start_row)
      if (show.sig.stars) {
        note_rows <- c(note_rows, note_start_row + 1)
      }
      if (!is.null(cor.methods) && length(cor.methods) > 0) {
        methods_row <- ifelse(show.sig.stars, note_start_row + 2, note_start_row + 1)
        note_rows <- c(note_rows, methods_row)
      }

      # Apply note styling only to the rows that actually contain notes
      # and only to the first column since notes are text that spans left
      for (row in note_rows) {
        openxlsx::addStyle(wb, sheet_name, note_style,
                           rows = row,
                           cols = 1)
      }

      # Auto-size columns
      openxlsx::setColWidths(wb, sheet_name, cols = 1:ncol(df_table), widths = "auto")

      # Save the workbook
      tryCatch({
        openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
        cat("Table saved to", filename, "\n")
      }, error = function(e) {
        warning("Error saving file: ", e$message)
      })
    }
  }

  return(result)
}

#' Print method for apa.cor.table objects
#' @param x An apa.cor.table object
#' @param ... Additional arguments (not used)
#' @export
print.apa.cor.table <- function(x, ...) {
  cat("\n", x$table.title, "\n\n")
  print(x$table.body, quote = FALSE)
  cat("\n", x$table.note, "\n")
  invisible(x)
}
