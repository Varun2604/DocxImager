
const {DocxImager} = require('./index');

(async ()=>{
    "use strict";
    let docxImager = new DocxImager();
    await docxImager.load("./test/insert_image_1.docx");

    await docxImager.insertImage({"img1" : "https://www.alambassociates.com/wp-content/uploads/2016/10/maxresdefault.jpg"});

    await docxImager.save("./test1.docx");
})();