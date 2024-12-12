#' Create Word Report
#'
#' This function generates a Word report that includes the input R code,
#' the execution results, and any plots produced.
#'
#' @param code A character string of R code to be executed.
#' @param output_path The file path where the Word document will be saved.
#' @return The path to the Word document.
#' @export
#'
#' @examples
#' code <- "
#'   library(ggplot2)
#'   p <- ggplot(mpg, aes(x = displ, y = hwy)) + geom_point()
#'   print(p)
#'   summary(mpg)
#' "
#' create_word(code, 'report.docx')
create_word <- function(code, output_path = "report.docx") {
  # 1. 將代碼中的單引號和雙引號處理為可兼容格式
  code <- gsub("'", "\"", code)  # 將單引號替換為雙引號

  # 2. 將代碼保存到臨時的 Rmarkdown 文件
  temp_rmd <- tempfile(fileext = ".Rmd")
  rmd_content <- paste0(
    "---\n",
    "title: \"R Code and Results Report\"\n",
    "output: word_document\n",
    "---\n\n",
    "```{r, echo=TRUE}\n",
    code,
    "\n```\n"
  )
  writeLines(rmd_content, temp_rmd)

  # 3. 使用 rmarkdown::render() 生成 Word 文件
  output_path <- normalizePath(output_path, mustWork = FALSE)
  rmarkdown::render(temp_rmd, output_file = output_path)

  message("Report saved at: ", output_path)
  return(output_path)
}
