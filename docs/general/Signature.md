#### **Signature**



代表附加到文档的数字签名。**Signature** 对象包含在 **Document** 对象的 **SignatureSet** 集合中。

**说明**

可以使用 **Add** 方法向 **SignatureSet** 集合中添加一个 **Signature** 对象，并使用 **Item** 方法返回现有的成员。要从 **SignatureSet** 集合中删除一个 **Signature** 对象，请使用 **Signature** 对象的 **Delete** 方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例提示用户选择要用于签署 WPS 中活动文档的数字签名。要使用此示例，请在 WPS 中打开一个文档，并向此函数传递与“数字证书”对话框中数字证书的“颁发者”和“颁发给”字段相匹配的证书颁发者名称和证书签署者名称。此示例将进行测试，以确保在将新签名提交到磁盘之前用户选择的数字签名满足特定条件，例如没有过期。*/ function test(strIssuer, strSigner) { 	var AddSignature;     try {         //Display the dialog box that lets the         //user select a digital signature.         //If the user selects a signature, then         //it is added to the Signatures         //collection. If the user does not, then         //an error is returned.         var sig = Application.ActiveWorkbook.Signatures.Add();          //Test several properties before commiting the Signature object to disk.         if(sig.Issuer == strIssuer && sig.Signer == strSigner && sig.IsCertificateExpired == false && sig.IsCertificateRevoked == false && sig.IsValid == true) {             alert("Signed");             AddSignature = true;         }         //Otherwise, remove the Signature object from the SignatureSet collection.         else {             sig.Delete();             alert("Not signed");             AddSignature = false ;         }          //Commit all signatures in the SignatureSet collection to the disk.         Application.ActiveWorkbook.Signatures.Commit();     }     catch(exception) {         AddSignature = false ;         alert("Action canceled.");     } } ` |

**方法**

|                                                              | 名称            | 说明                              |
| ------------------------------------------------------------ | --------------- | --------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**      | 从集合中删除 **Signature** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowDetails** | 显示与签名数据包相关的详细信息。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Sign**        | 创建一个签名数据包。              |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 获取一个 **Application** 对象，代表 **Signature** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanSetup**        | 获取一个 **Boolean** 类型的值，指示用户是否能设置 **Signature** 对象的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 获取一个 32 位整数，指示创建 **Signature** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Details**         | 获取有关签名的信息。只读。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsSignatureLine** | 获取一个指示此行是否为签名行的值。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsSigned**        | 获取一个 **Boolean** 类型的值，指示是否已成功签署文档。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 获取 Signature 对象的 **Parent** 对象。只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Setup**           | 获取 **SignatureSetup** 对象，该对象提供对签名数据包的各种属性的访问。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SignatureSetup**  | 获取与作为签名行的 **Signature** 对象相关联的 **Shape** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SortHint**        | 获取一个代表数据包（包含多个签名）中的签名排序顺序的值。只读。 |

**成员方法**

#### **Signature.Delete**

从集合中删除 **Signature** 对象。

**语法**

**express.Delete()**

*express*   一个代表 **Signature** 对象的变量。

**说明**

对于 **Scripts** 集合，使用 **Delete** 方法将删除指定的 WPS 文档、ET 工作表或 WPP 幻灯片中的所有脚本。在宿主应用程序中，脚本定位标记以形状表示。因此，与每个 **msoScriptAnchor** 类型的脚本定位标记相关联的 **Shape** 对象将从 ET 和 WPP 的 **Shapes** 集合以及 WPS 的 **InlineShapes** 和 **Shapes** 集合中删除。

#### **Signature.ShowDetails**

显示与签名数据包相关的详细信息。

**语法**

**express.ShowDetails()**

*express*   一个代表 **Signature** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将调用 **ShowDetails** 方法来显示 **Signature** 对象的详细信息。*/function test(objSignature) {     if(objSignature.IsSigned) {         alert("The document has been signed with the following details:" + objSignature.ShowDetails());     }else { 	alert("The document has not been signed.");     } }` |

#### **Signature.Sign**

创建一个签名数据包。

**语法**

**express.Sign(varSigImg, varDelSuggSigner, varDelSuggSignerLine2, varDelSuggSignerEmail)**

*express*   一个代表 **Signature** 对象的变量。

**参数**

| **名称**                | **必选/可选** | **数据类型** | **说明**                   |
| ----------------------- | ------------- | ------------ | -------------------------- |
| *varSigImg*             | 必选          | **Variant**  | 签名行的图形图像。         |
| *varDelSuggSigner*      | 必选          | **Variant**  | 建议签署人。               |
| *varDelSuggSignerLine2* | 必选          | **Variant**  | 附加签名行。               |
| *varDelSuggSignerEmail* | 必选          | **Variant**  | 建议签署人的电子邮件地址。 |

**说明**

调用 **Sign** 方法时，WPS Office将创建一个清单，并调用签名提供程序为文档中的每个流创建一个哈希。Office 随后会将结果捆绑到一个未签名的 XMLDSIG 模板中，并调用提供程序修改 XMLDSIG（如有必要），然后对其进行签署。生成的已签署签名之后将被退还给 Office 以便存储。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在下面的示例中，将设置代表签名图像、签署人、签署人职务以及签署人电子邮件的变量，然后将调用 Sign 方法来创建并签署签名数据包。*/ function test(sig, img){ 	var varSuggestedSigner = "Nancy Davolio"; 	var varSignatureTitle = "Sales Represenative"; 	var varSignerEmail = "ndavolio@northwindtraders.com"; 	sig.Sign(img, varSuggestedSigner, varSignatureTitle, varSignerEmail); }` |

**成员属性**

#### **Signature.Application**

获取一个 **Application** 对象，代表 **Signature** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.CanSetup**

获取一个 **Boolean** 类型的值，指示用户是否能设置 **Signature** 对象的属性。只读。

**语法**

**express.CanSetup**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.Creator**

获取一个 32 位整数，指示创建 **Signature** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.Details**

获取有关签名的信息。只读。

**语法**

**express.Details**

*express*   一个代表 **Signature** 对象的变量。

**说明**

可读取的信息包括：与签名关联的证书是否已到期、签名是否有效以及签名是否为只读。

#### **Signature.IsSignatureLine**

获取一个指示此行是否为签名行的值。只读。

**语法**

**express.IsSignatureLine**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.IsSigned**

获取一个 **Boolean** 类型的值，指示是否已成功签署文档。只读。

**语法**

**express.IsSigned**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.Parent**

获取 Signature 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.Setup**

获取 **SignatureSetup** 对象，该对象提供对签名数据包的各种属性的访问。只读。

**语法**

**express.Setup**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.SignatureSetup**

获取与作为签名行的 **Signature** 对象相关联的 **Shape** 对象。只读。

**语法**

**express.SignatureSetup**

*express*   一个代表 **Signature** 对象的变量。

#### **Signature.SortHint**

获取一个代表数据包（包含多个签名）中的签名排序顺序的值。只读。

**语法**

**express.SortHint**

*express*   一个代表 **Signature** 对象的变量。

适用环境：web

适用平台：windows/linux