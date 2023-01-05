#### **SignatureInfo**



代表用于创建数字签名或文档内签名的信息。

**方法**

|                                                              | 名称                                    | 说明                                                         |
| ------------------------------------------------------------ | --------------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetCertificateDetail**                | 显示与数字证书相关的指定详细信息。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetSignatureDetail**                  | 显示与签名相关的指定详细信息。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCertificateDetailByThumbprint** | 在通过个性特征对用户进行验证后，显示一个包含有关数字证书的信息的对话框。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectSignatureCertificate**          | 显示一个对话框，允许用户选择要用于签署文档的签名证书。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowSignatureCertificate**            |                                                              |

**属性**

|                                                              | 名称                               | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                    | 获取一个代表 **SignatureInfo** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CertificateVerificationResults** | 获取验证数字证书的结果。只读。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContentVerificationResults**     | 获取一个代表已签名文档的散列内容的验证结果的值。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                        | 获取一个 32 位整数，指示创建 **SignatureInfo** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsCertificateExpired**           | 获取一个 **Boolean** 类型的值，指示数字证书是否已到期。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsCertificateRevoked**           | 获取一个 **Boolean** 类型的值，指示数字证书是否已吊销。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsCertificateUntrusted**         | 获取一个 **Boolean** 类型的值，指示用于以数字方式签署文档的数字证书是否来自受信任的来源。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsValid**                        | 获取一个 **Boolean** 类型的值，指示在执行了签名验证过程后签名是否已成功获得验证。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**                       | 获取一个 **Boolean** 类型的值，指示 **SignatureInfo** 对象是否为只读。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SignatureComment**               | 获取或设置包含纳入签名数据包中的注释的值。可读/写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SignatureImage**                 | 获取或设置用于签署文档的图像的值。可读/写。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SignatureProvider**              | 获取一个标识所安装的签名提供程序加载项的值。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SignatureText**                  | 获取或设置用于签署此文档的签名文字的值。可读/写。            |

**成员方法**

#### **SignatureInfo.GetCertificateDetail**

显示与数字证书相关的指定详细信息。

**语法**

**express.GetCertificateDetail(certdet)**

*express*   一个代表 **SignatureInfo** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**          | **说明**                               |
| --------- | ------------- | --------------------- | -------------------------------------- |
| *certdet* | 必选          | **CertificateDetail** | 一个指定要显示的证书详细信息的枚举值。 |

**返回值**

Variant

**示例**

#### **SignatureInfo.GetSignatureDetail**

显示与签名相关的指定详细信息。

**语法**

**express.GetSignatureDetail(sigdet)**

*express*   一个代表 **SignatureInfo** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**        | **说明**                               |
| -------- | ------------- | ------------------- | -------------------------------------- |
| *sigdet* | 必选          | **SignatureDetail** | 一个指定要显示的签名详细信息的枚举值。 |

**返回值**

Variant

**示例**

#### **SignatureInfo.SelectCertificateDetailByThumbprint**

在通过个性特征对用户进行验证后，显示一个包含有关数字证书的信息的对话框。

**语法**

**express.SelectCertificateDetailByThumbprint(bstrThumbprint)**

*express*   一个代表 **SignatureInfo** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                               |
| ---------------- | ------------- | ------------ | -------------------------------------- |
| *bstrThumbprint* | 必选          | **String**   | 包含有关由个性特征标识的签署人的信息。 |

**示例**

#### **SignatureInfo.SelectSignatureCertificate**

显示一个对话框，允许用户选择要用于签署文档的签名证书。

**语法**

**express.SelectSignatureCertificate(ParentWindow)**

*express*   一个代表 **SignatureInfo** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**   | **说明**                           |
| -------------- | ------------- | -------------- | ---------------------------------- |
| *ParentWindow* | 必选          | **IOleWindow** | 包含证书选择对话框所在窗口的句柄。 |

**示例**

#### **SignatureInfo.ShowSignatureCertificate**

**语法**

**express.ShowSignatureCertificate(ParentWindow)**

*express*   一个代表 **SignatureInfo** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**   | **说明**                         |
| -------------- | ------------- | -------------- | -------------------------------- |
| *ParentWindow* | 必选          | **IOleWindow** | 包含“证书”对话框所在窗口的句柄。 |

**示例**

**成员属性**

#### **SignatureInfo.Application**

获取一个代表 **SignatureInfo** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.CertificateVerificationResults**

获取验证数字证书的结果。只读。

**语法**

**express.CertificateVerificationResults**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.ContentVerificationResults**

获取一个代表已签名文档的散列内容的验证结果的值。

**语法**

**express.ContentVerificationResults**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.Creator**

获取一个 32 位整数，指示创建 **SignatureInfo** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.IsCertificateExpired**

获取一个 **Boolean** 类型的值，指示数字证书是否已到期。只读。

**语法**

**express.IsCertificateExpired**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.IsCertificateRevoked**

获取一个 **Boolean** 类型的值，指示数字证书是否已吊销。只读。

**语法**

**express.IsCertificateRevoked**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.IsCertificateUntrusted**

获取一个 **Boolean** 类型的值，指示用于以数字方式签署文档的数字证书是否来自受信任的来源。只读。

**语法**

**express.IsCertificateUntrusted**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.IsValid**

获取一个 **Boolean** 类型的值，指示在执行了签名验证过程后签名是否已成功获得验证。只读。

**语法**

**express.IsValid**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.ReadOnly**

获取一个 **Boolean** 类型的值，指示 **SignatureInfo** 对象是否为只读。只读。

**语法**

**express.ReadOnly**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.SignatureComment**

获取或设置包含纳入签名数据包中的注释的值。可读/写。

**语法**

**express.SignatureComment**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.SignatureImage**

获取或设置用于签署文档的图像的值。可读/写。

**语法**

**express.SignatureImage**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.SignatureProvider**

获取一个标识所安装的签名提供程序加载项的值。只读。

**语法**

**express.SignatureProvider**

*express*   一个代表 **SignatureInfo** 对象的变量。

#### **SignatureInfo.SignatureText**

获取或设置用于签署此文档的签名文字的值。可读/写。

**语法**

**express.SignatureText**

*express*   一个代表 **SignatureInfo** 对象的变量。

适用环境：web

适用平台：windows/linux