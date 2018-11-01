
const fs = require('fs');
const https = require('https');

const JSZip = require('jszip');
const {parseString,Builder} = require('xml2js');
const uuid = require('uuid/v4');

const IMAGE_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const IMAGE_TYPE = 'image/png';

//1. replace image - the template image will be given in the docx, take the url/base64 for the new image in the argument, along with the image no.
//2. insert image - use replace tags {{insert_image `variable_name` height width}}, and take variable_name in the arguments.

class DocxImager {

    constructor(docx_path){
        this.__loadDocx(docx_path).catch((e)=>{
            console.log(e);
        });
    }

    async __loadDocx(docx_path){
        let zip = new JSZip();
        this.zip = await zip.loadAsync(fs.readFileSync(docx_path));
    }

    replaceWithImageURL(image_uri, image_id, type, cbk){
        this.__validateDocx();
        let req3 = https.request('https://upload.wikimedia.org/wikipedia/commons/b/bf/Test_card.png', (res) => {
            let buffer = [];
            res.on('data', (d) => {
                buffer.push(d);
            });
            res.on('end', ()=>{
                this.__replaceImage(Buffer.concat(buffer), image_id, type, cbk);
            });
        });

        req3.on('error', (e) => {
            console.error(e);
        });
        req3.end();
    }

    replaceWithLocalImage(image_path, image_id, type, cbk){
        this.__validateDocx();
        let image_buffer = fs.readFileSync(image_path);
        this.__replaceImage(image_buffer, image_id, type, cbk);
    }

    replaceWithB64Image(base64_string, image_id, type, cbk){
        this.__validateDocx();
        this.__replaceImage(Buffer.from(base64_string, 'base64'), image_id, type, cbk);
    }

    async __replaceImage(buffer, image_id, type, cbk){
        //1. replace the image
        let path = 'word/media/image'+image_id+'.'+type;
        this.zip.file(path, buffer);
        this.zip.generateNodeStream({streamFiles : true})
            .pipe(fs.createWriteStream('./merged.docx'))
            .on('finish', function(x){
                cbk();
            });
    }

    // {{insert_image variable_name type width height }} + {variable_name : "image_url"}
    //context - dict of variable_name vs url/b64 data
    async __insertImage(context, width, height, type, callback){
        //1. insert entry in [Content-Type].xml
        await this._addContentType(type);

        //2. write image in media folder in word/
        let image_path = await this._addImage(buffer);
        //3. insert entry in /word/_rels/document.xml.rels
        //<Relationship Id="rId3" Type=IMAGE_URI Target="media/image2.png"/>
        let rel_id = await this._addRelationship(image_path);
        //4. insert in document.xml after calculating EMU.
        await this._addInDocumentXML(rel_id, height, width);
        // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
        // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/

    }

    async _addContentType() {
        return new Promise(async (res, rej)=>{
            try{
                let content = await this.docx.file('[Content_Types].xml').async('nodebuffer');
                let matches = content.match(/<Types.*?>.*/gm);
                if (matches && matches[0]) {
                    let new_str = matches[0] + '<Default Extension="' + type + '" ContentType="image/' + type + '"/>'
                    content = content.replace(matches[0], new_str);

                    this.docx.file('[Content_Types].xml', content);
                    res(true);
                }
            }catch(e){
                console.log(e);
                rej(e);
            }
        })
    }

    async _addImage(image_buffer){
        return new Promise(async (res, rej)=>{
            try{
                let image_name = uuid();
                let image_path = 'media/'+image_name;
                this.docx.file('word/'+image_path, image_buffer);
                res(image_path);
            }catch(e){
                console.log(e);
                rej(e);
            }
        })
    }

    async _addRelationship(image_path){
        return new Promise(
            async function(res, rej){

                try{
                    let content = await this.docx.file('word/_rels/document.xml.rels').async('nodebuffer');
                    parseString(content.toString(), function(err, relation){
                        if(err){
                            console.log(err);       //TODO check if an error thrown will be catched by enclosed try catch
                            rej(err);
                            return;
                        }
                        let cnt = relation.Relationships.Relationship.length;
                        let rID = 'rId'+(cnt+1);
                        relation.Relationships.Relationship.push({
                            $ : {
                                Id : rID,
                                Type : IMAGE_URI,
                                Target : image_path
                            }
                        });
                        let builder = new Builder();
                        let modifiedXML = builder.buildObject(relation);
                        docx.file('word/_rels/document.xml.rels', modifiedXML);
                        res(rID);
                    });
                }catch(e){
                    console.log(e);
                    rej(e);
                }
            });
    }

    async _addInDocumentXML(rId, height, width){

        let calc_height = 9525 * height;
        let calc_width = 9525 * width;

        var xml_ele =
                '<w:rPr>' +
                    '<w:noProof/>' +
                '</w:rPr>' +
                '<w:drawing>' +
                    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
                        '<wp:extent cx="'+calc_width+'" cy="'+calc_height+'"/>' +
                        '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
                        '<wp:docPr id="1402" name="Picture" descr=""/>' +
                        '<wp:cNvGraphicFramePr>' +
                            '<a:graphicFrameLocks noChangeAspect="1"/>' +
                        '</wp:cNvGraphicFramePr>' +
                        '<a:graphic>' +
                            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
                                '<pic:pic>' +
                                    '<pic:nvPicPr>' +
                                        '<pic:cNvPr id="1" name="Picture" descr=""/>' +
                                        '<pic:cNvPicPr>' +
                                            '<a:picLocks noChangeAspect="0" noChangeArrowheads="1"/>' +
                                        '</pic:cNvPicPr>' +
                                    '</pic:nvPicPr>' +
                                    '<pic:blipFill>' +
                                        '<a:blip r:embed="'+rId+'">' +
                                            '<a:extLst>' +
                                                '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C2}">' +
                                                    '<a14:useLocalDpi val="0"/>' +
                                                '</a:ext>' +
                                            '</a:extLst>' +
                                        '</a:blip>' +
                                        '<a:srcRect/>' +
                                        '<a:stretch>' +
                                            '<a:fillRect/>' +
                                        '</a:stretch>' +
                                    '</pic:blipFill>' +
                                    '<pic:spPr bwMode="auto">' +
                                        '<a:xfrm>' +
                                            '<a:off x="0" y="0"/>' +
                                            '<a:ext cx="'+calc_width+'" cy="'+calc_height+'"/>' +
                                        '</a:xfrm>' +
                                        '<a:prstGeom prst="rect">' +
                                            '<a:avLst/>' +
                                        '</a:prstGeom>' +
                                        '<a:noFill/>' +
                                        '<a:ln>' +
                                            '<a:noFill/>' +
                                        '</a:ln>' +
                                    '</pic:spPr>' +
                                '</pic:pic>' +
                            '</a:graphicData>' +
                        '</a:graphic>' +
                    '</wp:inline>' +
                '</w:drawing>';

        return new Promise(async (res, rej)=>{
            try{
                let content = await this.docx.file('word/document.xml').async('nodebuffer');
                let matches = content.match(/(<w:p>.*?insert_image.*?<\/w:p>)/g);         //match all r tags
                if(matches && matches[0]){
                    let tag = matches[0].matches(/{{(.*?)}}/g)[0];
                    tag = tag.replace('<\/w:t>.*?<w:t>', '');
                    tag = tag.replace('\'', '');
                    let splits = tag.split(' ');
                    let href = splits[0];
                    let width = splits[1];
                    let height = splits[2];
                    res(true);
                }else{
                    rej(new Error('Invalid Docx'));
                }
            }catch(e){
                console.log(e);
                rej(e);
            }
        });

    }

    __getXMLElement(){

    }

    __validateDocx(){
        if(!this.zip){
            throw new Error('Invalid docx path or format. Please reinitialise instance.')
        }
    }
}

module.exports = DocxImager;