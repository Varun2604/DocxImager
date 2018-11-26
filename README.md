# DOCX IMAGER

A simple package with which you can dynamically insert images to your docx.

## Getting Started

```
npm i docximager --save
```

## What can DocxImager do for you ?

With DocxImager, you can either replace a template image in your document, or insert an image 
dynamically.

## What is replacing a template image ?

There are certain cases where you have to insert images to your docx document dynamically. You will need to place the dynamic image with certain formatting in your document to attain perfection.
Replacing template image functionality of DocxImager will come in handy at such places.

[!replace_docx_image](https://i.ibb.co/Q8YHdtW/replace-image-Google-Docs.png)

#### How to replace your images dynamically ?

Just follow these simple steps:

1. Include the package.
   ```
   const DocxImager = require('DocxImager');
   ```
   
2. Create a new Instance of DocxImager.
   ```
   let docxImager = new DocxImager();
   ```
   
3. Load your docx file.
   ```
   await docxImager.load('./my.docx')
   ```
   
4. Replace your image
   ```
   await docxImager.replaceWithImageURL('https://path_of_image.com/image.png', 1, 'png')
   ```
   
5. Save your new document in the specified path.
   ```
   await docxImager.save('./my_new_docx.docx');
   ```

## What is inserting image ?

There are certain cases where you will need to just insert an image, and formatting does not exactly matter. 
Such cases can be handeled with insert image functionality of DocxImager.

#### How to replace your images dynamically ?

Just follow these simple steps:

1. Include the package.
   ```
   const DocxImager = require('DocxImager');
   ```
   
2. Create a new Instance of DocxImager.
   ```
   let docxImager = new DocxImager();
   ```
   
3. Load your docx file.
   ```
   await docxImager.load('./my.docx')
   ```
   
4. Insert your image
   ```
   await docxImager.insertImage({'img1' : "https://path_of_image.com/image.png"})
   ```
   
5. Save your new document in the specified path.
   ```
   await docxImager.save('./my_new_docx.docx');
   ```


## Authors

* **Varun Venketeswaran Iyer** - *Initial work* - [Varun2604](https://github.com/Varun2604)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details



    