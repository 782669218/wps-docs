#### **SignatureSet**



对应于附加到文档的数字签名的 **Signature** 对象的集合。

**说明**

使用 **Document** 对象的 **Signatures** 属性可返回 **SignatureSet** 集合，例如：

| JavaScript for Application 复制                    |
| -------------------------------------------------- |
| `let sigs = Application.ActiveDocument.Signatures` |

可以使用 **Add** 方法向 **SignatureSet** 集合中添加一个 **Signature** 对象，并使用 **Item** 方法返回现有的成员。**AddSignatureLine** 方法也可向该集合中添加一个 **Signature** 对象。另请参见 **Subset** 属性，它可充当一个筛选器，用于决定该集合中是否出现某些 **Signature** 对象。要从 **SignatureSet** 集合中删除一个 **Signature** 对象，请使用 **Signature** 对象的 **Delete** 方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例提示用户选择要用于签署 WPS 中活动文档的数字签名。要使用此示例，请在 WPS 中打开一个文档，并向此函数传递与“数字证书”对话框中数字证书的“颁发者”和“颁发给”字段相匹配的证书颁发者名称和证书签署者名称。此示例将进行测试，以确保在将新签名提交到磁盘之前用户选择的数字签名满足特定条件，例如没有过期。*/ function test(strIssuer, strSigner) {     //Display the dialog box that lets the     //user select a digital signature.     //If the user selects a signature, then     //it is added to the Signatures     //collection. If the user doesn't, then     //an error is returned.     var sig = Application.ActiveWorkbook.Signatures.Add();      //Test several properties before committing the Signature object to disk.     if(sig.Issuer == strIssuer && sig.IsCertificateExpired == false && sig.IsCertificateRevoked == false && sig.IsValid == true) {       alert("Signed");       AddSignature = true;     }     //Otherwise, remove the Signature object from the SignatureSet collection.     else {         sig.Delete();         alert("Not signed");         AddSignature = false ;     } }` |

**方法**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddNonVisibleSignature** | 在以数字方式签署文档时创建一个签名数据包。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddSignatureLine**       | 向在其中收集签名的文档中添加行。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                   | 获取一个与文档当前所签署的数字签名之一相对应的 **Signature** 对象。只读。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 获取一个 **Application** 对象，代表 **SignatureSet** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanAddSignatureLine** | 获取一个 **Boolean** 类型的值，指示是否可向文档中添加签名行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**               | 获取一个 **Long** 类型的值，指示 **SignatureSet** 对象中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 获取一个 32 位整数，指示创建 **SignatureSet** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 获取 **SignatureSet** 对象的 **Parent** 对象。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSignaturesPane**  | 获取或设置一个 **Boolean** 类型的值，指示是否应显示**“签名”**任务窗格。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subset**              | 获取或设置一个可充当文档可用 **Signature** 对象上的筛选器的值。可读/写。 |

**成员方法**

#### **SignatureSet.AddNonVisibleSignature**

在以数字方式签署文档时创建一个签名数据包。

**语法**

**express.AddNonVisibleSignature(varSigProv)**

*express*   一个代表 **SignatureSet** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                |
| ------------ | ------------- | ------------ | ----------------------- |
| *varSigProv* | 可选          | **Variant**  | 代表签名提供程序的 ID。 |

**返回值**

Signature

**说明**

要提供用于触发此方法的入口点，您需要创建包含签名提供程序加载项的用户界面。此入口点通常以菜单项的形式提供给用户。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在以数字方式签署文档时，以下函数使用签名提供程序 ID 参数来创建签名数据包。*/ function test(varSigProviderID) {     var objSignatureSet = Application.ActiveWorkbook.Signatures;     var objSignature = Application.ActiveWorkbook.Signatures.Item(1);         objSignature = objSignatureSet.AddNonVisibleSignature(varSigProviderID);     return objSignature; }` |

#### **SignatureSet.AddSignatureLine**

向在其中收集签名的文档中添加行。

**语法**

**express.AddSignatureLine(varSigProv)**

*express*   一个代表 **SignatureSet** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                |
| ------------ | ------------- | ------------ | ----------------------- |
| *varSigProv* | 可选          | **Variant**  | 代表签名提供程序的 ID。 |

**返回值**

Signature

**说明**

添加了行后，文档的作者可以添加必要的信息，以使每个签名行都显示应签名人员的姓名和（可选）职称。当用户打开文档时，WPS Office可识别出存在一个或多个签名行，但签名行为空白。随后，Office 将向用户发出警告，告知他们需要签署此文档，并帮助用户找到所请求的签名在文档中的位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中的过程将接收签名提供程序的 ID，并添加签名行（只要文档不为只读）。*/ function test(SignProviderID) {     if(Application.ActiveWorkbook.Signatures.CanAddSignatureLine) { 	var objSignature = Application.ActiveWorkbook.Signatures.AddSignatureLine(SignProviderID);      }     return objSignature; }` |

#### **SignatureSet.Item**

获取一个与文档当前所签署的数字签名之一相对应的 **Signature** 对象。只读。

**语法**

**express.Item(iSig)**

*express*   一个代表 **SignatureSet** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                      |
| -------- | ------------- | ------------ | ----------------------------- |
| *iSig*   | 必选          | **Long**     | 确定要返回的 Signature 对象。 |

**成员属性**

#### **SignatureSet.Application**

获取一个 **Application** 对象，代表 **SignatureSet** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.CanAddSignatureLine**

获取一个 **Boolean** 类型的值，指示是否可向文档中添加签名行。只读。

**语法**

**express.CanAddSignatureLine**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.Count**

获取一个 **Long** 类型的值，指示 **SignatureSet** 对象中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.Creator**

获取一个 32 位整数，指示创建 **SignatureSet** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.Parent**

获取 **SignatureSet** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.ShowSignaturesPane**

获取或设置一个 **Boolean** 类型的值，指示是否应显示**“签名”**任务窗格。可读/写。

**语法**

**express.ShowSignaturesPane**

*express*   一个代表 **SignatureSet** 对象的变量。

#### **SignatureSet.Subset**

获取或设置一个可充当文档可用 **Signature** 对象上的筛选器的值。可读/写。

**语法**

**express.Subset**

*express*   一个代表 **SignatureSet** 对象的变量。

适用环境：web

适用平台：windows/linux