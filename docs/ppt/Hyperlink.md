**Hyperlink**



代表与非占位符形状或文本相关联的超链接。

**说明**

可以使用超链接跳转到一个 Internet 或 Intranet 网站、另一个文件或当前演示文稿中的一张幻灯片上。**Hyperlink** 对象是 **Hyperlinks** 集合的成员。**Hyperlinks** 集合包含幻灯片或母版中的所有超链接。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddToFavorites**    | 在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateNewDocument** | 新建一个与指定超链接相关联的 Web 演示文稿。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Follow**            | 在新的 Web 浏览器窗口中显示与指定超链接相关联的 HTML 文档。  |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Address**       | 返回或设置目标文档的 Internet 地址 (URL)。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailSubject**  | 返回或设置超链接主题行的文本字符串。该主题行附加到超链接的 Internet 地址 (URL) 中。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScreenTip**     | 返回或设置超链接的屏幕提示文字。可读/写 **String** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAndReturn** | 决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SubAddress**    | 返回或设置与指定超链接相关联的文档中的位置，例如 WPS 文档中的书签、ET 工作表中的区域或WPP 演示文稿中的幻灯片。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextToDisplay** | 返回或设置不是与图形相关联的超链接的显示文字。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**          | 返回一个 **MsoHyperlinkType** 常量，该常量代表超链接的类型。只读。 |

**成员方法**

#### **Hyperlink.AddToFavorites**

在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。

**语法**

**express.AddToFavorites()**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果文档具有显示名称，则快捷方式的名称是该显示名称；否则，快捷方式名称由 HLINK.DLL 决定。

**示例**

本示例在 Windows 程序文件夹的 Favorites 中添加一个指向当前演示文稿的超链接。

| 示例代码复制                                      |
| ------------------------------------------------- |
| `Application.ActivePresentation.AddToFavorites()` |

#### **Hyperlink.CreateNewDocument**

新建一个与指定超链接相关联的 Web 演示文稿。

**语法**

**express.CreateNewDocument(FileName, EditNow, Overwrite)**

*express*   一个代表 **Hyperlink** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**    | **说明**                                     |
| ----------- | ------------- | --------------- | -------------------------------------------- |
| *FileName*  | 必选          | **String**      | 文档的路径和文件名。                         |
| *EditNow*   | 必选          | **MsoTriState** | 确定是否在关联的编辑器中立即打开文档。       |
| *Overwrite* | 必选          | **MsoTriState** | 确定是否覆盖同一文件夹中现有的任何同名文件。 |

**返回值**

Nothing

**示例**

本示例将创建一个新的 Web 演示文稿，该演示文稿将与第一张幻灯片的第一个超链接相关联。新演示文稿的名为“Brittany.ppt”，同时它还会覆盖“HTMLPres”文件夹中的同名文件。此演示文稿还会立即装入到 WPP 中进行编辑。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).CreateNewDocument("C:\\HTMLPres\\Brittany.ppt", msoTrue, msoTrue)` |

#### **Hyperlink.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

#### **Hyperlink.Follow**

在新的 Web 浏览器窗口中显示与指定超链接相关联的 HTML 文档。

**语法**

**express.Follow()**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

本示例在一个新的 Web 浏览器实例中加载与第一张幻灯片的第一个超链接相关联的文档。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).Follow()` |

**成员属性**

#### **Hyperlink.Address**

返回或设置目标文档的 Internet 地址 (URL)。可读/写 **String** 类型。

**语法**

**express.Address**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例浏览第一张幻灯片中的所有形状，以查找指向 网站的 URL。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　for(let i=1;i<=myDocument.Hyperlinks.count;i++) { 　　　　    if(myDocument.Hyperlinks.Item(i).Address == "http://www.wps.cn/") {  　　　　       alert("You have a link to the  Home Page") 　　　　    } 　　　　} }` |

#### **Hyperlink.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Hyperlink.EmailSubject**

返回或设置超链接主题行的文本字符串。该主题行附加到超链接的 Internet 地址 (URL) 中。可读/写 **String** 类型。

**语法**

**express.EmailSubject**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

此属性通常与电子邮件超链接一起使用。其值优先于在同一 **Hyperlink** 对象的 **Address** 属性中指定的任意电子邮件主题。

**示例**

本示例设置活动演示文稿中第一张幻灯片上第一个超链接的电子邮件主题行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).EmailSubject ` |

#### **Hyperlink.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 **Shape** 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **Hyperlink.ScreenTip**

返回或设置超链接的屏幕提示文字。可读/写 **String** 类型。

**语法**

**express.ScreenTip**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

将演示文稿保存为 HTML、在 Web 浏览器中查看 Web 演示文稿以及将鼠标指针停留在超链接之上时，会显示屏幕提示文字。有些浏览器可能不支持屏幕提示。

**示例**

本示例为第一个超链接设置屏幕提示文字。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).ScreenTip = "Go to the WPS home page"` |

#### **Hyperlink.ShowAndReturn**

决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。**MsoTriState** 类型。

**语法**

**express.ShowAndReturn**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| msoCTrue                                                     |
| msoFalse 默认值。WPP不会从未激活的自定义幻灯片放映返回初始幻灯片放映。 |
| msoTriStateMixed                                             |
| msoTriStateToggle                                            |
| msoTrue：WPP从取消激活的自定义幻灯片放映返回初始幻灯片放映，该自定义幻灯片放映通过初始演示文稿的 Hyperlink 或 ActionSetting 对象激活。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置当前演示文稿中第一张幻灯片上第五个形状的鼠标单击动作为：显示名为“techtalk”的自定义幻灯片放映。当该自定义幻灯片放映结束后，将自动返回演示文稿在鼠标单击前的最初状态。*/ function test(){ 　　　　let as = Application.ActivePresentation.Slides.Item(1).Shapes.Item(5).ActionSettings.Item(ppMouseClick) 　　　　as.Action = ppActionNamedSlideShow 　　　　as.SlideShowName = "techtalk" 　　　　as.ShowandReturn = msoTrue }` |

#### **Hyperlink.SubAddress**

返回或设置与指定超链接相关联的文档中的位置，例如 WPS 文档中的书签、ET 工作表中的区域或WPP 演示文稿中的幻灯片。可读/写 **String** 类型。

**语法**

**express.SubAddress**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置在幻灯片放映过程中，单击当前演示文稿第一张幻灯片第一个形状时，跳转到 Latest Figures.ppt 中名为“Last Quarter”的幻灯片。*/ function test(){ 　　　　Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Action = ppActionHyperlink 　　　　let h = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Hyperlink 　　　　h.Address = "c:\\sales\\latest figures.ppt" 　　　　h.SubAddress = "last quarter" }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置在幻灯片放映过程中，单击当前演示文稿第一张幻灯片第一个形状时，跳转到 Latest.xls 中 A1:B10 单元格区域。*/ function test(){ 　　　　ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Action = ppActionHyperlink 　　　　let h = ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Hyperlink 　　　　h.Address = "c:\\sales\\latest.xls" 　　　　h.SubAddress = "A1:B10" }` |

#### **Hyperlink.TextToDisplay**

返回或设置不是与图形相关联的超链接的显示文字。可读/写 **String** 类型。

**语法**

**express.TextToDisplay**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果对不是与文字范围相关联的超链接使用此属性，则会产生运行时错误。可以使用以下类似代码检测给定的超链接（此处由 `myHyperlink` 代表）是否与某个文字范围相关联。

`If TypeName(myHyperlink.Parent.Parent) = "TextRange" Then strTRtest = "True" End If`

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个与第一张幻灯片第二个形状内的文字相关联的超链接。然后将显示文字设置为“WPS Home Page”，并将超链接地址设置为相应的 URL*/ function test(){ 　　　　let t = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.ActionSettings.Item(ppMouseClick) 　　　　t.Action = ppActionHyperlink 　　　　t.Hyperlink.TextToDisplay = "WPS Home Page" 　　　　t.Hyperlink.Address = "http://www.wps.cn" }` |

#### **Hyperlink.Type**

返回一个 **MsoHyperlinkType** 常量，该常量代表超链接的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

| MsoHyperlinkType 可以是下列 MsoHyperlinkType 常量之一。 |
| ------------------------------------------------------- |
| msoHyperlinkInlineShape 仅用于 WPS。                    |
| msoHyperlinkRange                                       |
| msoHyperlinkShape                                       |

适用环境：web

适用平台：windows/linux