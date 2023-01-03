**Shape**



代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

**说明**

| 注释                                                         |
| ------------------------------------------------------------ |
| 共有三个代表形状的对象：**Shapes** 集合，代表文档中的所有形状；[**ShapeRange**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange) 集合，代表文档中指定的形状子集（例如，**ShapeRange** 对象可以代表文档中的第一个和第四个形状，或代表文档中的所有选定形状）；**Shape** 对象，代表文档中的单个形状。若要同时使用多个形状或使用选定范围中的多个形状，请使用 **ShapeRange** 集合。有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。 |

以下示例说明如何执行下列操作：

- 按名称或编号索引，返回幻灯片上现有的形状。
- 返回幻灯片上新建的形状。
- 返回选定范围中的形状。
- 返回幻灯片上的幻灯片标题和其他占位符。
- 返回与连接符的端点相连的形状。
- 返回演示文稿的默认形状。
- 返回新建的任意多边形。
- 返回组合内的单个形状。
- 返回新组合的一组形状。

## 示例

使用 **Shapes**(*index*) 返回一个代表幻灯片上形状的 **Shape** 对象，其中 *index* 是形状名称或索引号。

向 

Shapes

 集合添加新的图形时，将对该新添加的图形赋以默认的名称。若要为图形指定更有意义的名称，可使用 

Name

 属性。下例向 myDocument 添加矩形，将其命名为“Red Square”，然后设置该矩形的前景色和线型。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let addshape = myDocument.Shapes.AddShape(msoShapeRectangle, 144, 144, 72, 72)     addshape.Name = "Red Square"     addshape.Fill.ForeColor.RGB = (255, 0, 0)     addshape.Line.DashStyle = msoLineDashDot     } ` |

若要在幻灯片中添加形状并返回一个代表新建形状的 **Shape** 对象，请使用 **Shapes** 集合的下列方法之一：[AddConnector](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddConnector)、[AddLine ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddLine)、[**AddMediaObject** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddMediaObject)、[AddOLEObject ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddOLEObject)、[AddPicture](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddPicture) 、[AddPlaceholder ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddPlaceholder)、[**AddShape** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddShape)、[AddTable ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddTable)、[AddTextbox ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddTextbox)、[AddTextEffect](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.AddTextEffect)  。

使用 **Selection.ShapeRange**(*index*) 返回一个代表选定范围中形状的 **Shape** 对象，其中 *index* 是形状名称或索引号。以下示例设置活动窗口选定范围中第一个形状的填充（假定选定范围中至少有一个形状）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWindow.Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)` |

使用 **Shapes.Title** 返回代表幻灯片标题的 **Shape** 对象。使用 **Shapes.AddTitle** 在无标题的幻灯片中添加标题并返回代表新建标题的 **Shape** 对象。使用 **Shapes.Placeholders**(*index*) 返回一个代表占位符的 **Shape** 对象，其中 *index* 是占位符的索引号。如果没有改变过幻灯片中形状的排列顺序，则以下三个语句是等价的（假设第一张幻灯片有标题）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test()  {     Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Font.Italic = true     Application.ActivePresentation.Slides.Item(1).Shapes.Placeholders.Item(1).TextFrame.TextRange.Font.Italic = true     Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Font.Italic = true }` |

若要返回代表演示文稿默认形状的 

Shape

 对象，请使用 

DefaultShape

 属性。

若要返回一个 

Shape

 对象，该对象代表连接符所连接的形状之一，请使用 

BeginConnectedShape

 或 

EndConnectedShape

 属性。

使用 **BuildFreeform** 和 **AddNodes** 方法定义新任意多边形的几何形状，使用 **ConvertToShape** 方法创建任意多边形并返回代表该形状的 **Shape** 对象。

使用 **GroupItems**(*index*) 返回代表组合形状中的单个形状的 **Shape** 对象，其中 *index* 是形状的名称或组合中的索引号。

使用 **Group** 或 **Regroup** 方法可对某一范围的形状进行组合，并返回代表新组合的单个 **Shape** 对象。形成一个组合后，可以像使用任何其他形状一样使用该组合。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**             | 适用于指定的形状格式，该格式已使用 **PickUp** 方法复制。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**              | 将指定对象复制到剪贴板。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**               | 删除指定对象并将其放到剪贴板中。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Flip**              | 绕水平或垂直轴翻转指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementLeft**     | 以指定点数水平移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementRotation** | 按指定度数改变指定形状绕 z 轴的旋转量。使用 **Rotation** 属性设置形状的绝对旋转量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementTop**      | 以指定点数垂直移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickUp**            | 复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleHeight**       | 以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleWidth**        | 根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**            | 选择指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**            | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimationSettings**   | 返回一个 **AnimationSettings** 对象，该对象代表所有可应用于指定形状的动画的特殊效果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShapeType**       | 返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**     | 设置或返回指定对象的背景样式。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**               | 返回当前 Shape 对象的 Chart 对象。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Child**               | 如果该形状是子形状，或者如果形状区域内的所有形状都是同一个父形状的子形状，则属性值为 **MsoTrue**。只读。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConnectionSiteCount** | 返回指定形状中的连结点的数量。只读。**Number** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                | 返回一个 **FillFormat** 对象，该对象包含指定形状的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                | 返回指定形状的发光格式。只读 **msoGlowType** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**          | 返回一个 **GroupShapes** 对象，该对象代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Item** 方法可返回形状组中的单个形状。适用于代表组合形状的 **Shape** 或 **ShapeRange** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChart**            | 返回指定对象代表的形状是否包含图表。只读。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTable**            | 返回指定的形状是否为表。只读。**MsoTriState** 类型。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTextFrame**        | 返回指定形状是否有文本框。只读。**MsoTriState** 类型。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Number** 类型；用于所有其他对象时可读/写，**Number**类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                  | 返回一个 **Number** 类型值，该值标识形状或形状范围。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                | 返回或设置一个 **Single** 类型值，该值代表从形状边界框的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MediaType**           | 返回 OLE 媒体类型。只读。**PpMediaType** 类型。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。**String** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**       | 返回一个 **PictureFormat** 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 **Shape** 或 **ShapeRange** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlaceholderFormat**   | 返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**              | 返回一个只读的 **ShadowFormat** 对象，该对象包含指定形状的阴影格式属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeStyle**          | 设置或返回指定对象的形状样式索引。可读/写。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Table**               | 返回一个 **Table** 对象，该对象代表某个形状或形状区域内的一个表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**          | 返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 **Shape** 或 **ShapeRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame**           | 返回一个 **TextFrame** 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**          | 返回与指定的 Shape 对象（包含指定形状的对齐方式和定位属性）相关联的 TextFrame2 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**              | 返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                 | 返回或设置一个 **Number**类型值，该值代表形状边界框的上边缘到文档上边缘的距离。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                | 返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**             | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读，**Number** 类型；用于所有其他对象时可读/写，**Number** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrderPosition**      | 返回指定的形状在 z-顺序中的位置。只读。**Number** 类型。     |

**成员方法**

#### **Shape.Apply**

适用于指定的形状格式，该格式已使用 **PickUp** 方法复制。

**语法**

**express.Apply()**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.Item(1).PickUp()     myDocument.Shapes.Item(2).Apply() }` |

#### **Shape.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。*/ /*如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败。*/ function test()  {     Application.Windows.Item(1).Selection.Copy()     Application.Windows.Item(2).View.Paste() }  /*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片。*/ function test()  {     let active = Application.ActivePresentation     active.Slides.Item(1).Shapes.Range([1, 2]).Copy()     active.Slides.Item(2).Shapes.Paste() }  /*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Copy()  /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()` |

#### **Shape.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/ Application.Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/ Application.ActivePresentation.Slides.Item(1).Cut()  /*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()` |

#### **Shape.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例从活动演示文稿的第一张幻灯片中删除所有任意多边形。*/ function test()  {     let shapes = Application.ActivePresentation.Slides.Item(1).Shapes     for(let i = shapes.Count; i >= 1; i--)     {         if(shapes.Item(i).Type == msoFreeform)         {             shapes.Item(i).Delete()         }     } }` |

#### **Shape.Flip**

绕水平或垂直轴翻转指定形状。

**语法**

**express.Flip(FlipCmd)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**   | **说明**                         |
| --------- | ------------- | -------------- | -------------------------------- |
| *FlipCmd* | 必选          | **MsoFlipCmd** | 指定形状是水平翻转还是垂直翻转。 |

**说明**

| MsoFlipCmd 可以是下列 MsoFlipCmd 常量之一。 |
| ------------------------------------------- |
| **msoFlipHorizontal**                       |
| **msoFlipVertical**                         |

#### **Shape.IncrementLeft**

以指定点数水平移动指定形状。

**语法**

**express.IncrementLeft(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Number**   | 以点为单位指定形状水平移动的距离。正值表示将形状向右移动，负值表示将形状向左移动。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let duplicate = myDocument.Shapes(1)     duplicate.Fill.PresetTextured(msoTextureGranite)     duplicate.IncrementLeft(70)     duplicate.IncrementTop(-50)     duplicate.IncrementRotation(30) }` |

#### **Shape.IncrementRotation**

按指定度数改变指定形状绕 z 轴的旋转量。使用 **Rotation** 属性设置形状的绝对旋转量。

**语法**

**express.IncrementRotation(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Number**   | 以度为单位指定形状的水平旋转量。正值表示使形状按顺时针方向旋转；负值表示使形状按逆时针方向旋转。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 **IncrementRotationX** 方法或 **IncrementRotationY** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let duplicate = myDocument.Shapes(1)     duplicate.Fill.PresetTextured(msoTextureGranite)     duplicate.IncrementLeft(70)     duplicate.IncrementTop(-50)     duplicate.IncrementRotation(30) }` |

#### **Shape.IncrementTop**

以指定点数垂直移动指定形状。

**语法**

**express.IncrementTop(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Number**   | 以点为单位指定形状对象的垂直移动距离。正值表示将形状向下移动，负值表示将形状向上移动 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的第一个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let duplicate = myDocument.Shapes(1)     duplicate.Fill.PresetTextured(msoTextureGranite)     duplicate.IncrementLeft(70)     duplicate.IncrementTop(-50)     duplicate.IncrementRotation(30) }` |

#### **Shape.PickUp**

复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。

**语法**

**express.PickUp()**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.Item(1).PickUp()     myDocument.Shapes.Item(2).Apply() }` |

#### **Shape.ScaleHeight**

以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。

**语法**

**express.ScaleHeight(Factor, RelativeToOriginalSize, fScale)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Number**       | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 指定是否相对于形状的当前或原始尺寸来缩放形状。               |
| *fScale*                 | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**：相对于形状的当前尺寸缩放。                     |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 **msoTrue**。 |

| MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 |
| ----------------------------------------------- |
| **msoScaleFromBottomRight**                     |
| **msoScaleFromMiddle**                          |
| **msoScaleFromTopLeft**：默认值。               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let s = myDocument.Shapes     for (let i = 1; i <= s.Count; i++)     {         switch(s.Item(i).Type)         {             case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:                 s.Item(i).ScaleHeight (1.75, msoTrue)                 s.Item(i).ScaleWidth (1.75, msoTrue)                 break             default:                 s.Item(i).ScaleHeight (1.75, msoFalse)                 s.Item(i).ScaleWidth (1.75, msoFalse)         }     } }` |

#### **Shape.ScaleWidth**

根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。

**语法**

**express.ScaleWidth(Factor, RelativeToOriginalSize, fScale)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Number**       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 指定是否相对于形状的当前或原始尺寸来缩放形状。               |
| *fScale*                 | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**：相对于形状的当前尺寸缩放。                     |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 **msoTrue**。 |

| MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 |
| ----------------------------------------------- |
| **msoScaleFromBottomRight**                     |
| **msoScaleFromMiddle**                          |
| **msoScaleFromTopLeft**：默认值。               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let s = myDocument.Shapes     for (let i = 1; i <= s.Count; i++)     {         switch(s.Item(i).Type)         {             case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:                 s.Item(i).ScaleHeight (1.75, msoTrue)                 s.Item(i).ScaleWidth (1.75, msoTrue)                 break             default:                 s.Item(i).ScaleHeight (1.75, msoFalse)                 s.Item(i).ScaleWidth (1.75, msoFalse)         }     } }` |

#### **Shape.Select**

选择指定的对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**    | **说明**                                   |
| --------- | ------------- | --------------- | ------------------------------------------ |
| *Replace* | 可选          | **MsoTriState** | 指定该选择是否替换之前所做的任何一项选择。 |

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

| MsoTriState 可以是下列 MsoTriState 常量之一。          |
| ------------------------------------------------------ |
| **msoCTrue**                                           |
| **msoFalse** 将该选择添加到前一项选择中。              |
| **msoTriStateMixed**                                   |
| **msoTriStateToggle**                                  |
| **msoTrue** 默认值。该选择替换之前所做的任何一项选择。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择活动演示文稿第一张幻灯片上的第一个和第三个形状。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Select() /*本示例将活动演示文稿第一张幻灯片上的第二个和第四个形状添加到前一项选择中。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Range([2, 4]).Select(false)` |

#### **Shape.ZOrder**

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。

**语法**

**express.ZOrder(ZOrderCmd)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                 |
| ----------- | ------------- | ---------------- | ---------------------------------------- |
| *ZOrderCmd* | 必选          | **MsoZOrderCmd** | 指定针对其他形状将指定形状移动到的位置。 |

**说明**

| MsoZOrderCmd 可以是下列 MsoZOrderCmd 常量之一。 |
| ----------------------------------------------- |
| **msoBringForward**                             |
| **msoBringInFrontOfText** 仅用于 WPS。          |
| **msoBringToFront**                             |
| **msoSendBackward**                             |
| **msoSendBehindText** 仅用于 WPS。              |
| **msoSendToBack**                               |

可用 **ZOrderPosition** 属性确定形状在 z-顺序中的当前位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myaddshape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)     while(myaddshape.ZOrderPosition > 2)     {         myaddshape.ZOrder(msoSendBackward)     } }` |

**成员属性**

#### **Shape.AnimationSettings**

返回一个 **AnimationSettings** 对象，该对象代表所有可应用于指定形状的动画的特殊效果。

**语法**

**express.AnimationSettings**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在创建幻灯片时，当前演示文稿第二张幻灯片的第一个形状从左侧飞入。*/ function test()  {     let asets = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).AnimationSettings     asets.EntryEffect = ppEffectFlyFromLeft     asets.TextLevelEffect = ppAnimateByAllLevels }` |

#### **Shape.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test()  {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= shpOle.Count; i++)     {         if(shpOle.Item(i).Type == msoLinkedOLEObject)         {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Shape.AutoShapeType**

返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

*express*   一个代表 **Shape** 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **ConnectorFormat** 对象的 **Type** 属性可设置或返回连接符类型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 myDocument 中用 32 角星替换所有 16 角星。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let s = myDocument.Shapes     for(let i = 1; i <= s.Count; i++)     {         if(s.Item(i).AutoShapeType == msoShape16pointStar)         {             s.Item(i).AutoShapeType = msoShape32pointStar         }     } }` |

#### **Shape.BackgroundStyle**

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Chart**

返回当前 Shape 对象的 Chart 对象。只读。

**语法**

**express.Chart**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Child**

如果该形状是子形状，或者如果形状区域内的所有形状都是同一个父形状的子形状，则属性值为 **MsoTrue**。只读。**MsoTriState** 类型。

**语法**

**express.Child**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue** 不应用于此属性。                                |
| **msoFalse**：该形状不是子形状，或者形状区域内并非所有的子形状都属于同一个父形状。 |
| **msoTriStateMixed** 不应用于该属性。                        |
| **msoTriStateToggle** 不应用于此属性。                       |
| **msoTrue**：该形状是子形状，或者形状区域内所有的子形状都属于同一个父形状。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择了画布中的第一个形状，如果选定的形状是子形状，则用指定的颜色填充该形状。以下示例假设当前演示文稿中的第一个形状是包含多个形状的绘图画布。*/ function test() {     //Select the first shape in the drawing canvas     Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).CanvasItems.Item(1).Select()     //Fill selected shape if it is a child shape     let actsel = Application.ActiveWindow.Selection     if(actsel.ShapeRange.Child == true)     {         actsel.ShapeRange.Fill.ForeColor.RGB = 100, 0, 200     }     else     {         alert("This shape is not a child shape.")     } }` |

#### **Shape.ConnectionSiteCount**

返回指定形状中的连结点的数量。只读。**Number** 类型。

**语法**

**express.ConnectionSiteCount**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加两个矩形，并且用两个连接符连接这两个矩形。两个连接符的起点都连接到第一个矩形的第一个连接位置，两个连接符的终点分别连接到第二个矩形的第一个和最后一个连接位置。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let s = myDocument.Shapes     let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)     let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)     let lastsite = secondRect.ConnectionSiteCount     let myaddcor = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat     myaddcor.BeginConnect(firstRect, 1)     myaddcor.EndConnect(secondRect, 1)     myaddcor.BeginConnect(firstRect, 1)     myaddcor.EndConnect(secondRect, lastsite) }` |

#### **Shape.Fill**

返回一个 **FillFormat** 对象，该对象包含指定形状的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myshapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill     myshapes.ForeColor.RGB = 128, 0, 0     myshapes.BackColor.RGB = 170, 170, 170     myshapes.TwoColorGradient (msoGradientHorizontal, 1) }` |

#### **Shape.Glow**

返回指定形状的发光格式。只读 **msoGlowType** 类型。

**语法**

**express.Glow**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.GroupItems**

返回一个 **GroupShapes** 对象，该对象代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Item** 方法可返回形状组中的单个形状。适用于代表组合形状的 **Shape** 或 **ShapeRange** 对象。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myshapes = myDocument.Shapes     myshapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"     myshapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"     myshapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"     let mygroup = myshapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()     mygroup.Fill.PresetTextured (msoTextureBlueTissuePaper)     mygroup.GroupItems.Item(2).Fill.PresetTextured (msoTextureGreenMarble) }` |

#### **Shape.HasChart**

返回指定对象代表的形状是否包含图表。只读。

**语法**

**express.HasChart**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.HasTable**

返回指定的形状是否为表。只读。**MsoTriState** 类型。

**语法**

**express.HasTable**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**：指定的形状是表。                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查当前选定的形状是否为表格。如果是表格，此代码将第一列的宽度设置为 1 英寸（72 磅）。*/ function test()  {     let srange = Application.ActiveWindow.Selection.ShapeRange     if(srange.HasTable == msoTrue)     {         srange.Table.Columns.Item(1).Width = 72     } }` |

#### **Shape.HasTextFrame**

返回指定形状是否有文本框。只读。**MsoTriState** 类型。

**语法**

**express.HasTextFrame**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。   |
| ----------------------------------------------- |
| **msoCTrue**                                    |
| **msoFalse**                                    |
| **msoTriStateMixed**                            |
| **msoTriStateToggle**                           |
| **msoTrue**：指定形状有文本框，因此可包含文本。 |

#### **Shape.Height**

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Number** 类型；用于所有其他对象时可读/写，**Number**类型。

**语法**

**express.Height**

*express*   一个代表 **Shape** 对象的变量。

**说明**

一个 **Height**Shape 对象的 **Shape**Height 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

#### **Shape.Id**

返回一个 **Number** 类型值，该值标识形状或形状范围。只读。

**语法**

**express.Id**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向当前演示文稿添加一个新形状，再根据 ID 属性的值填充该形状。*/ function test() {     let myaddshape = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 100, 100, 100, 100)     switch (myaddshape.Id)     {         case myaddshape.Id >0 && myaddshape.Id <=500:             myaddshape.Fill.ForeColor.RGB = 255, 0, 0             break         case myaddshape.Id >500 && myaddshape.Id <=1000:             myaddshape.Fill.ForeColor.RGB = 255, 255, 0             break         case myaddshape.Id >1000 && myaddshape.Id <=1500:             myaddshape.Fill.ForeColor.RGB = 255, 0, 255             break         case myaddshape.Id >1500 && myaddshape.Id <=2000:             myaddshape.Fill.ForeColor.RGB = 0, 255, 0             break         case myaddshape.Id >2000 && myaddshape.Id <=2500:             myaddshape.Fill.ForeColor.RGB = 0, 255, 255             break         default:             myaddshape.Fill.ForeColor.RGB = 0, 0, 255      } }` |

#### **Shape.Left**

返回或设置一个 **Single** 类型值，该值代表从形状边界框的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。

**语法**

**express.Left**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.MediaType**

返回 OLE 媒体类型。只读。**PpMediaType** 类型。

**语法**

**express.MediaType**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| PpMediaType 可以是下列 PpMediaType 常量之一。 |
| --------------------------------------------- |
| **ppMediaTypeMixed**                          |
| **ppMediaTypeMovie**                          |
| **ppMediaTypeOther**                          |
| **ppMediaTypeSound**                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在放映幻灯片时，当前演示文稿内第一张幻灯片上的所有原始声音对象循环播放直到被手动停止。*/ function test()  {     let so = Application.ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= so.Count; i++)     {         if(so.Item(i).Type == msoMedia)         {             if(so.Item(i).MediaType == ppMediaTypeSound)             {                 so.Item(i).AnimationSettings.PlaySettings.LoopUntilStopped = true             }         }     }   }` |

#### **Shape.Name**

创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。**String** 类型，可读/写。

**语法**

**express.Name**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿第一张幻灯片上的第二个对象的名称设为“big triangle”。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).Name = "big triangle"  /*本示例为当前演示文稿第一张幻灯片中名为“big triangle”的形状设置填充颜色。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item("big triangle").Fill.ForeColor.RGB = 0, 0, 255` |

#### **Shape.PictureFormat**

返回一个 **PictureFormat** 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 **Shape** 或 **ShapeRange** 对象。只读。

**语法**

**express.PictureFormat**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let mypfmat = myDocument.Shapes.Item(1).PictureFormat     mypfmat.Brightness = 0.3     mypfmat.Contrast = 0.75  }` |

#### **Shape.PlaceholderFormat**

返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。

**语法**

**express.PlaceholderFormat**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Shadow**

返回一个只读的 **ShadowFormat** 对象，该对象包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿第一张幻灯片中添加一个有阴影的矩形。蓝色浮凸效果的阴影距该矩形的右边 3 磅，下边 2 磅。*/ function test()  {     let myShap = Application.ActivePresentation.Slides.Item(1).Shapes     let myshadow = myShap.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow     myshadow.Type = msoShadow17     myshadow.ForeColor.RGB = (0, 0, 128)     myshadow.OffsetX = 3     myshadow.OffsetY = 2 }` |

#### **Shape.ShapeStyle**

设置或返回指定对象的形状样式索引。可读/写。

**语法**

**express.ShapeStyle**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Table**

返回一个 **Table** 对象，该对象代表某个形状或形状区域内的一个表格。只读。

**语法**

**express.Table**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第二张幻灯片第五个形状中表格第一列的宽度设置为 80 磅。*/ Application.ActivePresentation.Slides.Item(2).Shapes.Item(5).Table.Columns.Item(1).Width = 80 ` |

#### **Shape.TextEffect**

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 **Shape** 或 **ShapeRange** 对象。

**语法**

**express.TextEffect**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.TextFrame**

返回一个 **TextFrame** 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。

**语法**

**express.TextFrame**

*express*   一个代表 **Shape** 对象的变量。

**说明**

使用 **TextFrame** 对象的 **TextRange** 属性可返回文本框中的文本。

在应用 **TextFrame** 属性之前，可使用 **HasTextFrame** 属性来确定形状中是否包含文本框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在 myDocument 中添加一个矩形，向该矩形中添加文本，然后设置文本框的上边距。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myshapes = myDocument.Shapes.AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame     myshapes.TextRange.Text = "Here is some test text"     myshapes.MarginTop = 10 }` |

#### **Shape.TextFrame2**

返回与指定的 Shape 对象（包含指定形状的对齐方式和定位属性）相关联的 TextFrame2 对象。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.ThreeD**

返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为应用于 myDocument 中第一个形状的三维效果设置深度、延伸色、延伸方向以及照明方向。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let mythreed = myDocument.Shapes.Item(1).ThreeD     mythreed.Visible = true     mythreed.Depth = 50     //RGB value for purple     mythreed.ExtrusionColor.RGB = (255, 100, 255)     mythreed.SetExtrusionDirection (msoExtrusionTop)     mythreed.PresetLightingDirection = msoLightingLeft }` |

#### **Shape.Top**

返回或设置一个 **Number**类型值，该值代表形状边界框的上边缘到文档上边缘的距离。可读/写。

**语法**

**express.Top**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Type**

返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| MsoShapeType 可以是下列 MsoShapeType 常量之一。 |
| ----------------------------------------------- |
| **msoAutoShape**                                |
| **msoCallout**                                  |
| **msoCanvas**                                   |
| **msoChart**                                    |
| **msoComment**                                  |
| **msoDiagram**                                  |
| **msoEmbeddedOLEObject**                        |
| **msoFormControl**                              |
| **msoFreeform**                                 |
| **msoGroup**                                    |
| **msoLine**                                     |
| **msoLinkedOLEObject**                          |
| **msoLinkedPicture**                            |
| **msoMedia**                                    |
| **msoOLEControlObject**                         |
| **msoPicture**                                  |
| **msoPlaceholder**                              |
| **msoScriptAnchor**                             |
| **msoShapeTypeMixed**                           |
| **msoTable**                                    |
| **msoTextBox**                                  |
| **msoTextEffect**                               |

#### **Shape.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **Shape** 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse
msoTriStateMixed
msoTriStateToggle
msoTrue：指定对象或对象格式是可见的。

#### **Shape.Width**

以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读，**Number** 类型；用于所有其他对象时可读/写，**Number** 类型。

**语法**

**express.Width**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.ZOrderPosition**

返回指定的形状在 z-顺序中的位置。只读。**Number** 类型。

**语法**

**express.ZOrderPosition**

*express*   一个代表 **Shape** 对象的变量。

**说明**

Shapes(1) 返回 z-顺序中的最后一个形状，Shapes(Shapes.Count) 返回 z-顺序中的第一个形状。

若要设置形状在 z-顺序中的位置，请用 **ZOrder** 方法。

形状在 z-顺序中的位置与 **Shapes** 集合中形状的索引号相对应。例如，如果幻灯片上有四个形状，表达式 myDocument.Shapes(1) 在 z-顺序的后面返回该形状，表达式 myDocument.Shapes(4) 在 z-顺序的前面返回该形状。

如向一个集合添加新图形，其默认的添加位置是 z-顺序中的前面位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上。*/ function test()  {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myaddshape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)     while (myaddshape.ZOrderPosition > 2)     {         myaddshape.ZOrder (msoSendBackward)     } }` |

适用环境：web

适用平台：windows/linux