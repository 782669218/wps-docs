**WebOptions**



包含工作簿级的属性，当以网页保存文档或打开网页时，ET 将使用这些属性。

**说明**

您可以在应用程序（全局）级或在工作簿级返回或设置属性。（请注意，根据保存工作簿时的属性值，属性值可能会因工作簿而异。工作簿级的属性设置会重写应用程序级的属性设置。应用程序级属性包含在 **DefaultWebOptions** 对象中。

**示例**

使用 **WebOptions** 属性可返回 **WebOptions** 对象。下例实现的功能是：检查是否允许 PNG（便携网络图形）作为图像格式，然后设置相应的 `strImageFileType` 变量。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let strImageFileType     let objAppWebOptions = Application.Workbooks.Item(1).WebOptions     if (objAppWebOptions.AllowPNG == true) {         strImageFileType = "PNG"     }     else {         strImageFileType = "JPG"     } }` |

**方法**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UseDefaultFolderSuffix** | 根据所选或安装的语言支持，将含指定文档的文件夹后缀设置为默认后缀。 |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowPNG**             | 如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 **True**。如果不允许将 PNG 作为输出格式使用，则为 **False**。默认值是 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**              | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DownloadComponents**   | 如果必要的“Microsoft Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 **True**。如果没有下载这些组件，则为 **False**。默认值是 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Encoding**             | table { word-break:break-all; }当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。**MsoEncoding** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FolderSuffix**         | 返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True**）时，ET 使用该后缀。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LocationOfComponents** | 返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“Microsoft Office Web 组件”。默认值是 Microsoft Office 的本地或网络安装路径。**String** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OrganizeInFolder**     | 如果为 **True**，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 **False**，则支持文件和网页将保存在同一文件夹中。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PixelsPerInch**        | 返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelyOnCSS**            | 如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 **True**。ET 会创建一个级联样式表文件，然后依据 **OrganizeInFolder** 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML <FONT> 标记和级联样式表，则为 **False**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelyOnVML**            | 如果为 **True**，将文档保存为网页时，图形对象不能生成图像文件。如果为 **False**，则可生成图像文件。默认值是 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScreenSize**           | 返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 **MsoScreenSize** 常量之一。默认常量是 **msoScreenSize800x600**。**MsoScreenSize** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TargetBrowser**        | 返回或设置一个 **MsoTargetBrowser** 常量，它指明浏览器的版本。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseLongFileNames**     | 如果该属性值为 **True**，则将文档保存为网页时使用长文件名。如果该属性值为 **False**，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **WebOptions.UseDefaultFolderSuffix**

根据所选或安装的语言支持，将含指定文档的文件夹后缀设置为默认后缀。

**语法**

**express.UseDefaultFolderSuffix()**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

将文档保存为网页时，ET 会使用文件夹后缀，同时还会使用长文件名，并选择是否将支持文件保存在单独的文件夹中（即，是否将 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True**）。

此后缀位于文档名后面的文件夹名中。例如，如果文档名为“Book1”，语言为“英语”，则文件夹名为“Book1_files”。可用的文件夹后缀列在 **FolderSuffix** 属性主题中。

**示例**

本示例将第一个工作簿的文件夹后缀设置为默认后缀。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Workbooks.Item(1).WebOptions.UseDefaultFolderSuffix() }` |

**成员属性**

#### **WebOptions.AllowPNG**

如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 **True**。如果不允许将 PNG 作为输出格式使用，则为 **False**。默认值是 **False**。**Boolean** 类型，可读写。

**语法**

**express.AllowPNG**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果将图像保存为 PNG 格式而非其他文件格式，则可以提高图像的质量或减小图像文件的大小，因而也就可以减少下载所需的时间，前提是所使用的 Web 浏览器支持 PNG 格式。

**示例**

此示例使用 PNG 格式作为第一个工作簿的输出格式。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Workbooks.Item(1).WebOptions.AllowPNG = true }` |

或者，新建文档的应用程序可以将 PNG 作为全局默认设置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.DefaultWebOptions.AllowPNG = true } ` |

#### **WebOptions.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WebOptions** 对象的变量。

**示例**

本示例显示一条有关创建 `myObject` 的应用程序的消息。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myObject = Application.ActiveWorkbook     if (myObject.Application.Value == "ET") {         slert("This is an ET Application object.")     }     else {         slert("This is not an ET Application object.")     } }` |

#### **WebOptions.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **WebOptions.DownloadComponents**

如果必要的“Microsoft Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 **True**。如果没有下载这些组件，则为 **False**。默认值是 **False**。**Boolean** 类型，可读写。

**语法**

**express.DownloadComponents**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

可将 **LocationOfComponents** 属性设置为主 URL（在 Intranet 或网站上）或路径（本地或网络），授权用户在查看已保存的文档时，可从该路径所指向的位置下载组件。该路径必须是有效路径，且指向的位置必须包含必要的组件，同时用户还必须拥有有效的 Microsoft Office 许可。

Office Web 组件可增强保存为网页的文档的交互性。如果在没有安装这些组件的计算机上使用 Web 浏览器查看网页，则该页上的交互部分将表现为静态的（不能交互）。

**示例**

如果还没有安装“Office Web 组件”，此示例允许“Office Web 组件”与指定的网页一起下载。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.DefaultWebOptions.DownloadComponents = true     Application.DefaultWebOptions.LocationOfComponents = Application.Path + Application.PathSeparator + "foo" }` |

#### **WebOptions.Encoding**

table { word-break:break-all; }

当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。**MsoEncoding** 类型，可读写。

**语法**

**express.Encoding**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

不能使用任何带有 **AutoDetect** 后缀的常量。这些常量由 **ReloadAs** 方法使用。

#### **WebOptions.FolderSuffix**

返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True**）时，ET 使用该后缀。**String** 型，只读。

**语法**

**express.FolderSuffix**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

新建的文档将使用 **DefaultWebOptions** 对象的 **FolderSuffix** 属性所返回的后缀。如果文档曾在其他语言版本的 ET 中进行过编辑，则 **WebOptions** 对象的 **FolderSuffix** 属性的值可能会与 **DefaultWebOptions** 对象中该属性的值不同。可以使用 **UseDefaultFolderSuffix** 方法将后缀更改为当前在 Microsoft Office 中所使用的语言。

默认情况下，支持文件夹名由网页的名称加上一个下划线 (_)、一个句点 (.) 或连字号 (-)，再加上单词“files”（将文件保存为网页的 ET 版本的语言形式）组成。例如，如果使用荷兰语版本的 ET 将名为“Page1”的文件保存为网页，则支持文件夹的默认名称为“Page1_bestanden”。

下表列出 Office 的各个语言版本及相应的 **LanguageID** 属性值和文件夹后缀。对于表中没有列出的语言，其后缀为“.files”。

| 语言                 | LanguageID | 文件夹后缀   |
| -------------------- | ---------- | ------------ |
| 阿拉伯语             | 1025       | .files       |
| 巴斯克语             | 1069       | _fitxategiak |
| 巴西葡萄牙语         | 1046       | _arquivos    |
| 保加利亚语           | 1026       | .files       |
| 加泰罗尼亚语         | 1027       | _fitxers     |
| 中文 - 简体          | 2052       | .files       |
| 中文 - 繁体          | 1028       | .files       |
| 克罗地亚语           | 1050       | _datoteke    |
| 捷克语               | 1029       | _soubory     |
| 丹麦语               | 1030       | -filer       |
| 荷兰语               | 1043       | _bestanden   |
| 英语                 | 1033       | _files       |
| 爱沙尼亚语           | 1061       | _failid      |
| 芬兰语               | 1035       | _tiedostot   |
| 法语                 | 1036       | _fichiers    |
| 德语                 | 1031       | -Dateien     |
| 希腊语               | 1032       | .files       |
| 希伯来语             | 1037       | .files       |
| 匈牙利语             | 1038       | _elemei      |
| 意大利语             | 1040       | -file        |
| 日语                 | 1041       | .files       |
| 朝鲜语               | 1042       | .files       |
| 拉脱维亚语           | 1062       | _fails       |
| 立陶宛语             | 1063       | _bylos       |
| 挪威语               | 1044       | -filer       |
| 波兰语               | 1045       | _pliki       |
| 葡萄牙语             | 2070       | _ficheiros   |
| 罗马尼亚语           | 1048       | .files       |
| 俄语                 | 1049       | .files       |
| 塞尔维亚语(西里尔文) | 3098       | .files       |
| 塞尔维亚语（拉丁文） | 2074       | _fajlovi     |
| 斯洛伐克语           | 1051       | .files       |
| 斯洛文尼亚语         | 1060       | _datoteke    |
| 西班牙语             | 3082       | _archivos    |
| 瑞典语               | 1053       | -filer       |
| 泰语                 | 1054       | .files       |
| 土耳其语             | 1055       | _dosyalar    |
| 乌克兰语             | 1058       | .files       |
| 越南语               | 1066       | .files       |

**示例**

此示例返回第一个工作簿所使用的文件夹后缀。该后缀将返回到字符串变量 `strFolderSuffix` 中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let strFolderSuffix = Application.Workbooks.Item(1).WebOptions.FolderSuffix }` |

#### **WebOptions.LocationOfComponents**

返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“Microsoft Office Web 组件”。默认值是 Microsoft Office 的本地或网络安装路径。**String** 型，可读写。

**语法**

**express.LocationOfComponents**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果 **DownloadComponents** 属性设置为 **True**，Office Web 组件尚未安装，路径是有效路径且指向包含必要组件的位置，同时用户还具有有效的 Microsoft Office 许可，则会从指定网页自动下载这些组件。

#### **WebOptions.OrganizeInFolder**

如果为 **True**，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 **False**，则支持文件和网页将保存在同一文件夹中。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.OrganizeInFolder**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

在保存网页的文件夹中创建新文件夹，并依据文档命名。如果使用长文件名，则在文件夹名称后加后缀。**FolderSuffix** 属性返回已选择或已安装的语言支持的文件夹后缀，或返回默认的文件夹后缀。

如果在上次保存某个文档时，**OrganizeInFolder** 属性的值与现在不同，那么现在保存该文档时，ET 会相应地将支持文件自动移进或移出该文件夹。

如果并未使用长文件名（即，如果 **UseLongFileNames** 属性设置为 **False**），则 ET 会自动将任何支持文件保存到其他文件夹中。这些文件不能保存到与网页相同的文件夹中。

**示例**

此示例指定在将文档保存为网页时，所有的支持文件和网页保存在同一文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.DefaultWebOptions.OrganizeInFolder = false }` |

#### **WebOptions.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **WebOptions** 对象的变量。

#### **WebOptions.PixelsPerInch**

返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。**Long** 型，可读写。

**语法**

**express.PixelsPerInch**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

无论何时在 Web 浏览器中查看所保存文档，此属性都可确定指定的网页中相对于文本尺寸的图像以及单元格的大小。图像或单元格的实际最终尺寸等于原始尺寸（以英寸为单位）乘以每英寸中的像素数。

使用 **ScreenSize** 属性可为目标 Web 浏览器设置最佳屏幕大小。

**示例**

此示例根据浏览器的目标屏幕大小来设置像素密度。对于 800x600 像素的屏幕，其密度为 72 像素/英寸。对于 1024x768 像素的屏幕，其密度为 96 像素/英寸。对于所有其他情况，则使用的密度为 120 像素/英寸。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let web = Application.ActiveWorkbook.WebOptions     switch (web.ScreenSize) {         case msoScreenSize800x600:             web.PixelsPerInch = 72             break         case msoScreenSize1024x768:             web.PixelsPerInch = 96             break         default:             web.PixelsPerInch = 120     } }` |

#### **WebOptions.RelyOnCSS**

如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 **True**。ET 会创建一个级联样式表文件，然后依据 **OrganizeInFolder** 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML <FONT> 标记和级联样式表，则为 **False**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.RelyOnCSS**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果 Web 浏览器支持级联样式表，则应将此属性设置为 **True**，这样，就可以为该网页上提供更精确的布局和格式控制，网页也会更接近原来的文档（如同在 ET 中显示一样）。

**示例**

此示例启用级联样式表作为应用程序的全局默认值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveWorkbook.WebOptions.RelyOnCSS = true }` |

#### **WebOptions.RelyOnVML**

如果为 **True**，将文档保存为网页时，图形对象不能生成图像文件。如果为 **False**，则可生成图像文件。默认值是 **False**。**Boolean** 类型，可读写。

**语法**

**express.RelyOnVML**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果您的 Web 浏览器支持 Vector Markup Language (VML)，也可不从图形对象生成图像，这样可减少文件大小。例如，Microsoft Internet Explorer 5 支持此项功能，因此如果要将其作为目标浏览器，应该将 **RelyOnVML** 属性设为 **True**。如果浏览器不支持 VML，则在查看用此属性保存的网页时，将不会显示图像。

例如，如果网页中使用了已生成的图像文件，并且文档保存的位置与该网页在 Web 服务器上的最终位置不同，则不应生成图像。

**示例**

此示例指定在将工作表保存到网页上时，生成图像。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Workbooks.Item(1).WebOptions.RelyOnVML = false }` |

#### **WebOptions.ScreenSize**

返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 **MsoScreenSize** 常量之一。默认常量是 **msoScreenSize800x600**。**MsoScreenSize** 类型，可读写。

**语法**

**express.ScreenSize**

*express*   一个代表 **WebOptions** 对象的变量。

#### **WebOptions.TargetBrowser**

返回或设置一个 **MsoTargetBrowser** 常量，它指明浏览器的版本。可读写。

**语法**

**express.TargetBrowser**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

| **MsoTargetBrowser** 可为以下这些 **MsoTargetBrowser** 常量之一。 |
| ------------------------------------------------------------ |
| **msoTargetBrowserIE4** 。Microsoft Internet Explorer 4.0 或更高版本。 |
| **msoTargetBrowserIE5** 。Microsoft Internet Explorer 5.0 或更高版本。 |
| **msoTargetBrowserIE6** 。Microsoft Internet Explorer 6.0 或更高版本。 |
| **msoTargetBrowserV3** 。Microsoft Internet Explorer 3.0、Netscape Navigator 3.0 或更高版本。 |
| **msoTargetBrowserV4** 。Microsoft Internet Explorer 4.0、Netscape Navigator 4.0 或更高版本。 |

**示例**

在本示例中，ET 确定 Web 选项的浏览器版本是否为 IE5，并通知用户。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let wkbOne = Application.Workbooks.Item(1)      //Determine if IE5 is the target browser.     if (wkbOne.WebOptions.TargetBrowser == Application.Enum.msoTargetBrowserIE5) {         alert("The target browser is IE5 or later.")     } else {         alert("The target browser is not IE5 or later.")     } }` |

#### **WebOptions.UseLongFileNames**

如果该属性值为 **True**，则将文档保存为网页时使用长文件名。如果该属性值为 **False**，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.UseLongFileNames**

*express*   一个代表 **WebOptions** 对象的变量。

**说明**

如果不使用长文件名且文档还具有支持文件，ET 会自动将那些文件组织到单独的文件夹中。另外，使用 **OrganizeInFolder** 属性可确定支持文件是否组织在单独的文件夹中。

**示例**

此示例禁止将长文件名用作应用程序的全局默认设置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.DefaultWebOptions.UseLongFileNames = false }` |

适用环境：web

适用平台：windows/linux