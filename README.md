# PDF_Splitter

QP_410 Project requested in Qualtrics

A program used to prepare a large PDF file with multiple resumes into &lt;20Mb zip files while screening out files that are not resumes
Needs to be placed in zip files for no more than 20MB (Phenom People can automate the rest)

This program has no Qualtrics Specific information and can shared with no risk of exposing sensitive information.

pyInstaller would not compile the PyPdf2 module, so when compiling into an executable file, I have manually copied the file and imported a "custom" module that is the exact same and named it "PyPDF2_IMPORT"
