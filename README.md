# DOCX IMAGER

A simple package with which you can dynamically insert images to your docx.

## Getting Started

```
npm i docximager --save
```


## How to insert your images dynamically ?

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
   
4. Insert or replace your image
   ```
   await docxImager.replaceWithImageURL('https://path_of_inage.com/image.png')
   ```
   
5. Save your new document in the specified path.
   ```
   await docxImager.save('./my_new_docx.docx');
   ```

## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Authors

* **Varun Venketeswaran Iyer** - *Initial work* - [Varun2604](https://github.com/Varun2604)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details



    