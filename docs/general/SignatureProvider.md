#### **SignatureProvider**



代表签名提供程序加载项。

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GenerateSignatureLineImage** | 获取签名行图像。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetProviderDetail**          | 向签名提供程序加载项查询各种详细信息。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **HashStream**                 | 允许签名提供程序加载项为文档创建一个哈希值，您可以使用该值来确定在进行数字签名后文档内容是否被篡改过。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NotifySignatureAdded**       | 用于显示一个对话框，通知用户签署过程已完成并为加载项提供附加功能。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowSignatureDetails**       | 使签名提供程序加载项能够显示有关已签署签名行的详细信息，并显示诸如安全时间戳等其他已存储信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowSignatureSetup**         | 使签名提供程序加载项能够向用户显示**“签名设置”**对话框。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowSigningCeremony**        | 使签名提供程序加载项能够向用户显示**“签名”**对话框，从而允许用户指定他们的身份并随后获得验证。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SignXmlDsig**                | 用于签署 XMLDSIG 模板。                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **VerifyXmlDsig**              | 基于文档的签署状态和用于签名的证书的合法性来验证签名。       |

**成员方法**

#### **SignatureProvider.GenerateSignatureLineImage**

获取签名行图像。

**语法**

**express.GenerateSignatureLineImage(siglnimg, psigsetup, psiginfo)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**           | **说明**                             |
| ----------- | ------------- | ---------------------- | ------------------------------------ |
| *siglnimg*  | 必选          | **SignatureLineImage** | 如果为签名行图形，则包含相应的名称。 |
| *psigsetup* | 必选          | **SignatureSetup**     | 指定签名提供程序加载项的初始设置。   |
| *psiginfo*  | 必选          | **SignatureInfo**      | 指定有关签名提供程序加载项的信息。   |

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。对于在文档内容中显示的每个图像都调用此方法。可以异步调用该方法。例如，可以在签名设置之后直接为“未签名”图像和“无软件”图像调用该方法。然后可以在为“已签名”图像签名之后调用该方法。使用的四个图像是：

- **siglnimgSoftwareRequired**：如果在用户计算机上未安装签名提供程序加载项，则向用户显示此图像。如果用户尝试签名或查看签名行，则他们将被重定向到提供程序提供的超链接（在 **GetProviderDetail** 方法中指定）。
- **siglnimgUnsigned**：对于未签名的签名图像，将显示此图像。基本上，当加载具有未签名的签名行的文档时，签名提供程序将提示最新的签名图像，并显示此图像。
- **siglnimgSignedValid**：这是在签名行已进行签名且有效的情况下显示的图像（或者，更具体地说，已进行签名且签名没有注册为无效）。在打开文档时，假定在验证过程完成之前所有的已签署签名都是有效的，此时对于无效的签名显示“已签名/无效”图像。由于签名验证需要很长时间，因此签名验证在后台线程上与 Office 并行运行。由于加载项实现签名验证，因此代码将与 Office 并行运行，并且不应该试图在签名或验证期间显示 UI。
- **siglnimgSignedInvalid**：这是在对签名行进行了签名但签名存在问题（如文档已修改或者用户证书已作废）时显示的图像。由于加载项实现签名验证，因此您可以决定签名无效的方式和时间。

#### **SignatureProvider.GetProviderDetail**

向签名提供程序加载项查询各种详细信息。

**语法**

**express.GetProviderDetail(sigprovdet)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**                | **说明**                                                 |
| ------------ | ------------- | --------------------------- | -------------------------------------------------------- |
| *sigprovdet* | 必选          | **SignatureProviderDetail** | 包含一个枚举值，该枚举值代表要向加载项查询的信息的类型。 |

**返回值**

Variant

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。此方法用于向加载项查询三段信息：

- 加载项支持哪种哈希算法？
- 加载项是否只是用户界面 (UI)，或者它是否支持哈希处理和验证？如果返回 **TRUE**，则表明 WPS Office不会调用加载项进行哈希处理或验证，而只是调用加载项来显示 UI。
- 如果用户缺少签名加载项，加载项应向用户提供什么 URL？

#### **SignatureProvider.HashStream**

允许签名提供程序加载项为文档创建一个哈希值，您可以使用该值来确定在进行数字签名后文档内容是否被篡改过。

**语法**

**express.HashStream(QueryContinue, Stream)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**       | **说明**                                                     |
| --------------- | ------------- | ------------------ | ------------------------------------------------------------ |
| *QueryContinue* | 必选          | **IQueryContinue** | 提供了一种方式，可向主机应用程序进行查询以获得继续进行哈希处理的许可。 |
| *Stream*        | 必选          | **IStream**        | 包含数据流。                                                 |

**返回值**

Byte

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。将为文档中的每个签名数据流调用一次此方法。返回值是一个代表使用哈希算法计算出的哈希值的字节数组。

#### **SignatureProvider.NotifySignatureAdded**

用于显示一个对话框，通知用户签署过程已完成并为加载项提供附加功能。

**语法**

**express.NotifySignatureAdded(ParentWindow, psigsetup, psiginfo)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**       | **说明**                                           |
| -------------- | ------------- | ------------------ | -------------------------------------------------- |
| *ParentWindow* | 必选          | **IOleWindow**     | 允许主机应用程序获取包含所显示对话框的窗口的句柄。 |
| *psigsetup*    | 必选          | **SignatureSetup** | 包含签名提供程序的初始设置。                       |
| *psiginfo*     | 必选          | **SignatureInfo**  | 包含有关签名提供程序加载项的信息。                 |

**说明**

将在签署过程完成后调用此方法。使签名提供程序加载项能够为加载项增加附加功能。例如，如果想要提供用户可在其中上载已签署文档的存档服务，则可以使用此方法来启动该过程。

#### **SignatureProvider.ShowSignatureDetails**

使签名提供程序加载项能够显示有关已签署签名行的详细信息，并显示诸如安全时间戳等其他已存储信息。

**语法**

**express.ShowSignatureDetails(ParentWindow, psigsetup, psiginfo, XmlDsigStream, pcontverres, pcertverres)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**                       | **说明**                                   |
| --------------- | ------------- | ---------------------------------- | ------------------------------------------ |
| *ParentWindow*  | 必选          | **IOleWindow**                     | 包含窗口的句柄，该窗口中包含签名详细信息。 |
| *psigsetup*     | 必选          | **SignatureSetup**                 | 指定签名提供程序的初始设置。               |
| *psiginfo*      | 必选          | **SignatureInfo**                  | 指定有关已签署签名行的信息。               |
| *XmlDsigStream* | 必选          | **IStream**                        | 代表数据流或 XML 的二进制大对象。          |
| *pcontverres*   | 必选          | **ContentVerificationResults**     | 包含一个值，代表验证签名内容的结果。       |
| *pcertverres*   | 必选          | **CertificateVerificationResults** | 包含一个值，代表验证签名证书的结果。       |

#### **SignatureProvider.ShowSignatureSetup**

使签名提供程序加载项能够向用户显示**“签名设置”**对话框。

**语法**

**express.ShowSignatureSetup(ParentWindow, psigsetup)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**       | **说明**                                       |
| -------------- | ------------- | ------------------ | ---------------------------------------------- |
| *ParentWindow* | 必选          | **IOleWindow**     | 包含窗口的句柄，该窗口中包含“签名设置”对话框。 |
| *psigsetup*    | 必选          | **SignatureSetup** | 指定签名提供程序的初始设置。                   |

**说明**

在插入时配置过程中，以及在用户稍后想要重新配置签名行的情况下，都将使用此方法。将在此回调过程中显示**“签名设置”**对话框并等待用户选择**“确定”**或**“取消”**。除非明确需要作者提供有关签名行的信息，否则不必为签名设置显示对话框。如果可以将所有必要的详细信息提供给 WPS Office而无需用户输入，则不需要对话框。

#### **SignatureProvider.ShowSigningCeremony**

使签名提供程序加载项能够向用户显示**“签名”**对话框，从而允许用户指定他们的身份并随后获得验证。

**语法**

**express.ShowSigningCeremony(ParentWindow, psigsetup, psiginfo)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**       | **说明**                                   |
| -------------- | ------------- | ------------------ | ------------------------------------------ |
| *ParentWindow* | 必选          | **IOleWindow**     | 包含窗口的句柄，该窗口中包含“签名”对话框。 |
| *psigsetup*    | 必选          | **SignatureSetup** | 指定签名提供程序的初始设置。               |
| *psiginfo*     | 必选          | **SignatureInfo**  | 指定有关签名提供程序的信息。               |

**说明**

当用户尝试在签名行上签名时，或者，如果加载项已在 Office 应用程序的对象模型中对 **SignatureLine** 对象调用了 **Sign** 方法，WPS Office应用程序将在内部调用此方法。

#### **SignatureProvider.SignXmlDsig**

用于签署 XMLDSIG 模板。

**语法**

**express.SignXmlDsig(QueryContinue, psigsetup, psiginfo, XmlDsigStream)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**       | **说明**                                                     |
| --------------- | ------------- | ------------------ | ------------------------------------------------------------ |
| *QueryContinue* | 必选          | **IQueryContinue** | 提供了一种方式，可查询主机应用程序以获得继续进行验证操作的许可。 |
| *psigsetup*     | 必选          | **SignatureSetup** | 指定有关签名行的配置信息。                                   |
| *psiginfo*      | 必选          | **SignatureInfo**  | 指定从签署仪式中捕获的信息。                                 |
| *XmlDsigStream* | 必选          | **IStream**        | 代表一个包含 XML 的数据流，该数据流又代表一个 XMLDSIG 对象。 |

**说明**

XMLDSIG 是一种基于标准的签名格式 (http://www.w3.org/TR/xmldsig-core/)，可由第三方验证。此格式是 WPS Office中的默认签名格式。

#### **SignatureProvider.VerifyXmlDsig**

基于文档的签署状态和用于签名的证书的合法性来验证签名。

**语法**

**express.VerifyXmlDsig(QueryContinue, psigsetup, psiginfo, XmlDsigStream, pcontverres, pcertverres)**

*express*   一个代表 **SignatureProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**                       | **说明**                                                     |
| --------------- | ------------- | ---------------------------------- | ------------------------------------------------------------ |
| *QueryContinue* | 必选          | **IQueryContinue**                 | 提供了一种方式，可查询主机应用程序以获得继续进行验证操作的许可。 |
| *psigsetup*     | 必选          | **SignatureSetup**                 | 指定有关签名行的配置信息。                                   |
| *psiginfo*      | 必选          | **SignatureInfo**                  | 指定从签署仪式中捕获的信息。                                 |
| *XmlDsigStream* | 必选          | **IStream**                        | 代表一个包含 XML 的数据流，该数据流又代表一个 XMLDSIG 对象。 |
| *pcontverres*   | 必选          | **ContentVerificationResults**     | 指定签名验证操作的状态。                                     |
| *pcertverres*   | 必选          | **CertificateVerificationResults** | 指定签名证书验证的状态。                                     |

**说明**

XMLDSIG 是一种基于标准的签名格式 (http://www.w3.org/TR/xmldsig-core/)，可由第三方验证。此格式是 WPS Office中的默认签名格式。

适用环境：web

适用平台：windows/linux