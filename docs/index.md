<a name="DocxImager"></a>

## DocxImager
**Kind**: global class  

* [DocxImager](#DocxImager)
    * [new DocxImager()](#new_DocxImager_new)
    * [.load(docx_path)](#DocxImager+load) ⇒ <code>Promise</code>
    * [.replaceWithImageURL(image_uri, image_id, type)](#DocxImager+replaceWithImageURL) ⇒ <code>Promise</code>
    * [.replaceWithLocalImage(image_path, image_id, type)](#DocxImager+replaceWithLocalImage) ⇒ <code>Promise</code>
    * [.replaceWithB64Image(base64_string, image_id, type)](#DocxImager+replaceWithB64Image) ⇒ <code>Promise</code>
    * [.save(op_file_name)](#DocxImager+save) ⇒ <code>Promise</code>

<a name="new_DocxImager_new"></a>

### new DocxImager()
<p>Represents a DocxImager instance</p>

<a name="DocxImager+load"></a>

### docxImager.load(docx_path) ⇒ <code>Promise</code>
<p>Load the DocxImager instance with the docx.</p>

**Kind**: instance method of [<code>DocxImager</code>](#DocxImager)  

| Param | Type | Description |
| --- | --- | --- |
| docx_path | <code>String</code> | <p>full path of the template docx</p> |

<a name="DocxImager+replaceWithImageURL"></a>

### docxImager.replaceWithImageURL(image_uri, image_id, type) ⇒ <code>Promise</code>
<p>Replaces the template image with the image obtained from the web url</p>

**Kind**: instance method of [<code>DocxImager</code>](#DocxImager)  

| Param | Type | Description |
| --- | --- | --- |
| image_uri | <code>String</code> | <p>web uri of the image</p> |
| image_id | <code>String</code> | <p>id of the image in the docx</p> |
| type | <code>String</code> | <p>type of the template image</p> |

<a name="DocxImager+replaceWithLocalImage"></a>

### docxImager.replaceWithLocalImage(image_path, image_id, type) ⇒ <code>Promise</code>
<p>Replaces the template image with the image obtained from the local path</p>

**Kind**: instance method of [<code>DocxImager</code>](#DocxImager)  

| Param | Type | Description |
| --- | --- | --- |
| image_path | <code>String</code> | <p>full path of the image in the local system</p> |
| image_id | <code>String</code> | <p>id of the image in the docx</p> |
| type | <code>String</code> | <p>type of the template image</p> |

<a name="DocxImager+replaceWithB64Image"></a>

### docxImager.replaceWithB64Image(base64_string, image_id, type) ⇒ <code>Promise</code>
<p>Replaces the template image with the image obtained from the Base64 string</p>

**Kind**: instance method of [<code>DocxImager</code>](#DocxImager)  

| Param | Type | Description |
| --- | --- | --- |
| base64_string | <code>String</code> | <p>Base64 form of the image</p> |
| image_id | <code>String</code> | <p>id of the image in the docx</p> |
| type | <code>String</code> | <p>type of the template image</p> |

<a name="DocxImager+save"></a>

### docxImager.save(op_file_name) ⇒ <code>Promise</code>
<p>Saves the transformed docx.</p>

**Kind**: instance method of [<code>DocxImager</code>](#DocxImager)  

| Param | Type | Description |
| --- | --- | --- |
| op_file_name | <code>String</code> | <p>Output file name with full path.</p> |
