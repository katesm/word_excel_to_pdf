# Word and Excel to PDF

The project is a .net core web api service that uses the office to pdf executable to convert office docs to a pdf. 

https://github.com/cognidox/OfficeToPDF

## Endpoints

http:localhost:5000/wordToPdf

http:localhost:5000/excelToPdf

Payload POST data must be a json object that looks like below.
<code>
  
  { FileName : "name_of_word_document.docx', FileBytes: [] //Byte array of file }
  
</code>
  
