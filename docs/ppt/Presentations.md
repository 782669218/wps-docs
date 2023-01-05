**Presentations**



WPP中所有 [**Presentation**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation) 对象的集合。每个 **Presentation** 对象代表WPP 中当前打开的一个演示文稿。

**说明**

**Presentations** 集合不包含开放式加载宏，它是一种特殊类型的隐藏演示文稿。但是如果知道某单个开放式加载宏的文件名，就可以返回它。例如，Presentations("oscar.ppa") 将以 **Presentation** 对象的形式返回名为“Oscar.ppa”的开放式加载宏。但是，建议使用 **AddIns** 集合返回开放式加载宏。

示例：

使用 [**Presentations**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentations/Presentations%20.htm#jsObject_Presentations) 属性返回 **Presentation** 集合。使用 **Add** 方法创建一个新演示文稿并将其添加到集合中。以下示例创建一个新演示文稿，在其中添加一张幻灯片，然后保存该演示文稿。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let newPres = Application.Presentations.Add(true)     newPres.Slides.Add(1, 1)     newPres.SaveAs("Sample") } ` |

使用 **Presentations**(*index*) 返回单个 **Presentation** 对象，其中 *index* 是演示文稿的名称或索引号。以下示例打印第一个演示文稿。

| 示例代码复制                                   |
| ---------------------------------------------- |
| `Application.Presentations.Item(1).PrintOut()` |

使用 [**Open**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentations/Presentations%20.htm#Presentations.Open) 方法打开演示文稿并将其添加到 **Presentations** 集合中。以下示例以只读方式打开文件“sales.ppt”。

| 示例代码复制                                       |
| -------------------------------------------------- |
| `Application.Presentations.Open("sales.ppt",true)` |

**方法**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**         | 创建一个演示文稿。返回一个代表新演示文稿的 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanCheckOut** | 值为 **True** 时，WPP 可以将指定的演示文稿从服务器签出。**Boolean** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckOut**    | 从服务器向本地计算机复制指定演示文稿以进行编辑。返回表示签出演示文稿的本地路径和文件名的 **String** 类型的值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**        | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Open**        | 打开指定的演示文稿。返回一个 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象，该对象代表已打开的演示文稿。 |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **Presentations.Add**

创建一个演示文稿。返回一个代表新演示文稿的 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象。

**语法**

**express.Add(WithWindow)**

*express*   一个代表 **Presentations** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**           | **说明**                       |
| ------------ | ------------- | ---------------------- | ------------------------------ |
| *WithWindow* | 可选          | **MsoTriState 类型。** | 是否在可视窗口中创建演示文稿。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个演示文稿，并在其中添加一张幻灯片，然后保存该演示文稿。*/ function test() {     let add = Application.Presentations.Add()     add.Slides.Add(1, ppLayoutTitle)     add.SaveAs("Sample") }` |

#### **Presentations.CanCheckOut**

值为 **True** 时，WPP 可以将指定的演示文稿从服务器签出。**Boolean** 类型，可读/写。

**语法**

**express.CanCheckOut(FileName)**

*express*   一个代表 **Presentations** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                     |
| ---------- | ------------- | ------------ | ---------------------------- |
| *FileName* | 必选          | **String**   | 演示文稿的服务器路径和名称。 |

**说明**

若要利用WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例验证演示文稿是否可以签出且未被其他用户签出。如果演示文稿可以签出，则将演示文稿复制到本地计算机以进行编辑。 function CheckOutPresentation(strPresentation) {     if (Application.Presentations.CanCheckOut(strPresentation) == true) {         Application.Presentations.CheckOut(strPresentation)     }     else {         alert("You are unable to check out this " + "presentation at this time.")     } }` |

#### **Presentations.CheckOut**

从服务器向本地计算机复制指定演示文稿以进行编辑。返回表示签出演示文稿的本地路径和文件名的 **String** 类型的值。

**语法**

**express.CheckOut(FileName)**

*express*   一个代表 **Presentations** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                     |
| ---------- | ------------- | ------------ | ---------------------------- |
| *FileName* | 必选          | **String**   | 演示文稿的服务器路径和名称。 |

**说明**

若要利用 WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例验证演示文稿是否可以签出且未被其他用户签出。如果演示文稿可以签出，则将演示文稿复制到本地计算机以进行编辑。 function CheckOutPresentation(strPresentation) {     if (Application.Presentations.CanCheckOut(strPresentation) == true) {         Application.Presentations.CheckOut(strPresentation)     }     else {         alert("You are unable to check out this " + "presentation at this time.")     } }` |

#### **Presentations.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Presentations** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

#### **Presentations.Open**

打开指定的演示文稿。返回一个 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象，该对象代表已打开的演示文稿。

**语法**

**express.Open(FileName, ReadOnly, Untitled, WithWindow, OpenConflictDocument)**

*express*   一个代表 **Presentations** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**    | **说明**                                       |
| ---------------------- | ------------- | --------------- | ---------------------------------------------- |
| *FileName*             | 必选          | **String**      | 要打开的文件名。                               |
| *ReadOnly*             | 可选          | **MsoTriState** | 指定是以可读/写状态还是以只读状态打开文件。    |
| *Untitled*             | 可选          | **MsoTriState** | 指定文件是否有标题。                           |
| *WithWindow*           | 可选          | **MsoTriState** | 指定文件是否可见。                             |
| *OpenConflictDocument* | 可选          | **MsoTriState** | 指定是否为带有脱机冲突的演示文稿打开冲突文件。 |

**说明**

安装适当的文件转换程序后，WPP 可打开具有下列 MS-DOS 文件扩展名的文件：.ch3、.cht、.doc、.htm、.html、.mcw、.pot、.ppa、.pps、.ppt、.pre、.rtf、.sh3、.shw、.txt、.wk1、.wk3、.wk4、.wpd、.wpf、.wps 和 .xls。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse** 默认值。以可读/写状态打开文件。  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 以只读状态打开文件。              |

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 默认值。自动使用文件名作为被打开演示文稿的标题。 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 打开一个没有标题的文件。等于创建一个文件副本。   |

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse** 隐藏被打开的演示文稿。           |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 默认值。在可视窗口中打开文件。    |

| MsoTriState 可以是下列 MsoTriState 常量之一。       |
| --------------------------------------------------- |
| **msoCTrue**                                        |
| **msoFalse** 默认值。打开服务器文件并忽略冲突文档。 |
| **msoTriStateMixed**                                |
| **msoTriStateToggle**                               |
| **msoTrue** 打开冲突文件并覆盖服务器文件。          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以只读状态打开一个演示文稿。*/ Application.Presentations.Open("C:\\My Documents\\pres1.ppt", true)` |

**成员属性**

#### **Presentations.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Presentations** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中在活动的演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test()  {     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     for(let i=1; i<= Application.ActivePresentation.Slides.Item(1).Shapes.Count; i++)      {         let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes.Item(i)         if(shpOle.Type == msoLinkedOLEObject)          {             alert(shpOle.OLEFormat.Application.Name)         }     } }` |

#### **Presentations.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Presentations** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示所有打开的演示文稿的名字。*/ function test() {     let pres = Application.Presentations     for(let i=1; i<=pres.Count; i++)      {         alert(pres.Item(i).Name)     } }` |

#### **Presentations.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Presentations** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux