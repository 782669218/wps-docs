**ColorSchemes**



指定的演示文稿中所有 **ColorScheme** 对象的集合。每个 **ColorScheme** 对象代表一种配色方案，即同时用于幻灯片的一组颜色。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 将一种配色方案添加到可用方案集合中。返回一个代表所添加配色方案的 **ColorScheme** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从指定集合中返回单个对象。                                   |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **ColorSchemes.Add**

将一种配色方案添加到可用方案集合中。返回一个代表所添加配色方案的 **ColorScheme** 对象。

**语法**

**express.Add(Scheme )**

*express*   一个代表 **ColorSchemes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                                     |
| -------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Scheme* | 必选          | **ColorScheme** | ColorScheme 对象。要添加的配色方案。可以是任意幻灯片或母版中的 ColorScheme 对象，或者是任何打开的演示文稿的 ColorSchemes 集合中的一项。如果省略此参数，则使用所指定的演示文稿的 ColorScheme 集合中的第一个 ColorSchemes 对象（即第一个标准配色方案）。 |

**返回值**

ColorScheme

**说明**

新配色方案基于指定幻灯片或母版使用的颜色，或者一个打开的演示文稿中指定配色方案中的颜色。

**ColorSchemes** 集合最多可以包含 16 种配色方案。如果需要向集合中添加另一配色方案但 **ColorSchemes** 集合已满，可以使用 **Delete** 方法删除某个现有的配色方案。

请注意，WPP 在用户界面上添加配色方案之前会自动检测重复的配色方案；但是在 Visual Basic 的过程中添加配色方案时，不会进行这样的检测。过程必须进行自检以避免添加多余的配色方案。

#### **ColorSchemes.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ColorSchemes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号。 |

**返回值**

ColorScheme

**成员属性**

#### **ColorSchemes.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ColorSchemes** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **ColorSchemes.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ColorSchemes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭第二个窗口。*/ function test(){ 　　　　let w = Application.Windows     for(let i = 2; i <= w.Count; i++) { 　　　　    w.Item(2).Close() 　　　　} }` |

#### **ColorSchemes.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ColorSchemes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux