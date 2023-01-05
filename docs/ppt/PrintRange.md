**PrintRange**



代表要打印的连续幻灯片或页的范围。

**说明**

**PrintRange** 对象是 **PrintRanges** 集合的成员。**PrintRanges** 集合包含已为指定演示文稿定义的所有打印范围。

可以在 **PrintRanges** 集合中设置独立于 **RangeType** 设置的打印区域。这些打印区域在包含它们的演示文稿加载时始终有效。**RangeType** 属性设为 **ppPrintSlideRange** 时，应用 **PrintRanges** 集合中的区域。

**方法**

|                                                              | 名称       | 说明             |
| ------------------------------------------------------------ | ---------- | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **End**         | 返回指定打印范围中最后一张幻灯片的编号。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Start**       | 返回要打印的幻灯片范围内第一张幻灯片的编号。只读 **Integer** 类型。 |

**成员方法**

#### **PrintRange.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **PrintRange** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**成员属性**

#### **PrintRange.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **PrintRange** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") } ` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **PrintRange.End**

返回指定打印范围中最后一张幻灯片的编号。只读 **Long** 类型。

**语法**

**express.End**

*express*   一个代表 **PrintRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条信息，该信息指示当前演示文稿中第一个打印范围的起始和结束幻灯片的编号。*/ function test() {     let print = Application.ActivePresentation.PrintOptions.Ranges     if (print.Count > 0) {         let item = print.Item(1)         alert("Print range 1 starts on slide " + item.Start + " and ends on slide " + item.End)     } }` |

#### **PrintRange.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PrintRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **PrintRange.Start**

返回要打印的幻灯片范围内第一张幻灯片的编号。只读 **Integer** 类型。

**语法**

**express.Start**

*express*   一个代表 **PrintRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条信息，该信息指示当前演示文稿中第一个打印范围的起始和结束幻灯片的编号。*/ function test() {     let print = Application.ActivePresentation.PrintOptions.Ranges     if (print.Count > 0) {         let item = print.Item(1)         alert("Print range 1 starts on slide " + item.Start + " and ends on slide " + item.End)     } }` |

适用环境：web

适用平台：windows/linux