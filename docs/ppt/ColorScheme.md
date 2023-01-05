**ColorScheme**



代表一种配色方案，该配色方案是八种颜色的组合，分别用于幻灯片、备注页或讲义中的不同元素，例如标题或背景（请注意，演示文稿中幻灯片、备注页及讲义的配色方案可以单独设置）。

**说明**

每种颜色由一个 **RGBColor** 对象代表。**ColorScheme** 对象是 **ColorSchemes** 集合的成员。**ColorSchemes** 集合包含演示文稿中所有的配色方案。

以下示例说明如何执行下列操作：

- 从演示文稿的所有配色方案的集合中返回 **ColorScheme** 对象
- 返回附加到特定幻灯片或母版的 **ColorScheme** 对象
- 从 **ColorScheme** 对象返回单个幻灯片元素的颜色

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Colors** | 返回一个 **RGBColor** 对象，该对象代表配色方案中的一种颜色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的对象                                               |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **ColorScheme.Colors**

返回一个 **RGBColor** 对象，该对象代表配色方案中的一种颜色。

**语法**

**express.Colors(SchemeColor)**

*express*   一个代表 **ColorScheme** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**           | **说明**                     |
| ------------- | ------------- | ---------------------- | ---------------------------- |
| *SchemeColor* | 必选          | **PpColorSchemeIndex** | 指定的配色方案中的一种颜色。 |

**返回值**

RGBColor

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置当前演示文稿中第一张和第三张幻灯片的标题颜色。*/ function test() {     let mySlides = Application.ActivePresentation.Slides.Range([1, 3])     mySlides.ColorScheme.Colors(ppTitle).RGB = 0, 255, 0 }` |

#### **ColorScheme.Delete**

删除指定的对象

**语法**

**express.Delete()**

*express*   一个代表 **ColorScheme** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**成员属性**

#### **ColorScheme.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ColorScheme** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **ColorScheme.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ColorScheme** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭第二个窗口。*/ function test(){ 　　　　let w = Application.Windows     for(let i = 2; i <= w.Count; i++) { 　　　　    w.Item(2).Close() 　　　　} }` |

#### **ColorScheme.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ColorScheme** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux