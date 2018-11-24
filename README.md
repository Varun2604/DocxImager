A simple package with which you can dynamically insert images to your docx.

Just follow these simple steps:

1. Include the package.
   const DocxImager = require('DocxImager');
   
2. Create a new Instance of DocxImager.
   let dImager = new DocxImager();
   
3. Load your docx file.
   await dImager.load('./my.docx')
   
4. Insert or replace your image
   await dImager.replaceWithImageURL('https://path_of_inage.com/image.png')
   
5. Save your new document in the specified path.
   await dImager.save('./my_new_docx.docx');
    