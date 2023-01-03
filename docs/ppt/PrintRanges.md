**PrintRanges**



指定的演示文稿中所有 **PrintRange** 对象的集合。每个 **PrintRange** 对象代表一个要打印的连续幻灯片或页的范围。

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**      | 返回一个代表新动画行为的 **PrintRange** 对象。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearAll** | 从 **PrintRanges** 集合中清除所有打印范围。使用 **PrintRanges** 集合的 **Add** 方法可向该集合中添加打印范围。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**     | 从指定集合中返回单个对象。                                   |

**属性**

|                                                              | 名称            | 说明                                                      |
| ------------------------------------------------------------ | --------------- | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                    |

**成员方法**

#### **PrintRanges.Add**

返回一个代表新动画行为的 **PrintRange** 对象。

**语法**

**express.Add(Start, End)**

*express*   一个代表 **PrintRanges** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                       |
| -------- | ------------- | ------------ | ------------------------------ |
| *Start*  | 必选          | **Long**     | 范围中起始幻灯片的幻灯片编号。 |
| *End*    | 必选          | **Long**     | 范围中终止幻灯片的幻灯片编号。 |

**返回值**

PrintRange

#### **PrintRanges.ClearAll**

从 **PrintRanges** 集合中清除所有打印范围。使用 **PrintRanges** 集合的 **Add** 方法可向该集合中添加打印范围。

**语法**

**express.ClearAll()**

*express*   一个代表 **PrintRanges** 对象的变量。

**返回值**

Nothing

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除当前演示文稿中所有以前定义的打印范围；创建新的打印范围，其中包含第一张，第三张到第五张，第八张和第九张幻灯片；打印新定义的幻灯片范围；然后清除新定义的打印范围。*/ function test(){ 　　　　let print = ActivePresentation.PrintOptions     　　　　print.RangeType = ppPrintSlideRange    　　　　 let ranges = print.Ranges     　　　　    ranges.ClearAll()    　　　　     ranges.Add(1, 1)      　　　　   ranges.Add(3, 5)       　　　　  ranges.Add(8, 9)       　　　　  ranges.Parent.Parent.PrintOut()       　　　　  ranges.ClearAll() }` |

#### **PrintRanges.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **PrintRanges** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号。 |

**返回值**

PrintRange

**成员属性**

#### **PrintRanges.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者

**语法**

**express.Application**

*express*   一个代表 **PrintRanges** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") } ` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let slides = ActivePresentation.Slides.Item(1).Shapes        for(let i = 1; i <= slides.Count; i++){     　　　　if(slides.Item(i).Type == msoLinkedOLEObject){      　　　　   MsgBox(slides.Item(i).OLEFormat.Application.Name) 　　　　    } 　　　　} }` |

#### **PrintRanges.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **PrintRanges** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test(){ 　　　　let windows = Application.Windows 　　　　for(let i = 2; i <= windows.Count; i++){ 　　　　    window.Item(i).Close() 　　　　} }` |

#### **PrintRanges.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PrintRanges** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes 　　　　let addshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     　　　　addshape.TextRange.Text = "Test text"     　　　　addshape.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux