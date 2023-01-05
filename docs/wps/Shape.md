#### **Shape**



代表绘制层中的对象，如自选图形、任意多边形、OLE 对象、ActiveX 控件或图片。**Shape** 对象是 **Shapes** 集合的一个成员，该集合包含了某个文档的正文或个文档的所有页眉和页脚中的所有形状。

**说明**

形状总是属于某个锁定范围。可以将形状置于包含锁定标记的页面上的任何位置。有三种代表形状的对象：

1. **Shapes** 集合，代表文档中所有的形状
2. **ShapeRange** 对象，代表文档中指定的形状子集（例如，**ShapeRange** 对象可以代表文档中的形状一和形状四，也可以代表文档中的所有选定形状）；
3. **Shape** 对象，代表文档中的一个形状。如果需要同时处理若干个形状或处理所选内容中的若干个形状，可使用 **ShapeRange** 集合。

使用 **Shapes**(*Index*) 可返回一个 **Shape** 对象，其中 *Index* 为名称或索引号。以下示例水平翻转活动文档中的形状一。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Shapes.Item(1).Flip(msoFlipHorizontal)` |

以下示例水平翻转活动文档中名为“Rectangle 1”的形状。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Shapes.Item("Rectangle 1").Flip(msoFlipHorizontal)` |

在创建每个形状时为其指定一个默认名称。例如，如果向文档中添加了三个不同的形状，则这三个形状的名称可以是“矩形 2”、“文本框 3”和“椭圆 4”。要使形状的名称更有意义，请设置 **Name** 属性。

使用 **ShapeRange**(*Index*) 可返回代表所选内容中某个形状的 **Shape** 对象，其中 *Index* 为名称或索引号。假定所选内容包含至少一个形状，则以下示例为所选内容中的第一个形状设置填充效果。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)` |

假定所选内容至少包含一个形状，则以下示例为该区域的所有形状设置填充效果。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Selection.ShapeRange.Item(1).Fill.ForeColor.RGB = (255, 0, 0)` |

要将 

Shape

 对象添加到指定文档的形状集合，并返回一个代表新建形状的 

Shape

 对象，可使用 

Shapes

 集合的下列方法之一：

AddCallout

、

AddCurve

、

AddLabel

、

AddLine

、

AddOleControl

、

AddOleObject

、

AddPolyline

、

AddShape

、

AddTextbox

、

AddTextEffect

 或 

BuildFreeForm

。以下示例向活动文档添加一个矩形。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)` |

使用 **GroupItems**(*Index*) 可返回一个代表组合形状中的单个形状的 **Shape** 对象，其中 *Index* 为形状的名称或形状在该组中的索引号。

使用 **Group** 或 **Regroup** 方法可对一系列形状进行组合，并返回一个代表新组合的 **Shape** 对象。形成一个组合之后，便可以像处理任何其他形状一样处理该组合。

每个 **Shape** 对象都锁定到某个文本范围。将某个形状锁定到包含锁定范围的第一段的开头。该形状总是与其锁定标记在同一页上。

将 **ShowObjectAnchors** 属性设置为 **True** 可以查看锁定标记本身。形状的 **Top** 和 **Left** 属性决定了其垂直和水平位置。形状的 **RelativeHorizontalPosition** 和 **RelativeVerticalPosition** 属性决定了形状的位置是从锁定段落、包含锁定段落的栏、页边距还是页边缘开始计算。

如果将形状的 **LockAnchor** 属性设置为 **True**，则无法将锁定标记从其所在位置拖动到页面上。

使用 **Fill** 属性可返回 **FillFormat** 对象，该对象包含用于设置闭合形状的填充格式的所有属性和方法。**Shadow** 属性返回 **ShadowFormat** 对象，该对象用于设置阴影的格式。使用 **Line** 属性可返回 **LineFormat** 对象，该对象包含用于设置线条格式和箭头格式的各种属性和方法。**TextEffect** 属性返回 **TextEffectFormat** 对象，该对象用于设置艺术字的格式。**Callout** 属性返回 **CalloutFormat** 对象，该对象用于设置线形标注的格式。**WrapFormat** 属性返回 **WrapFormat** 对象，该对象用于定义文字环绕形状的方式。**ThreeD** 属性返回 **ThreeDFormat** 对象，该对象用于创建三维形状。可使用 **PickUp** 和 **Apply** 方法将一个形状的格式设置传送给另一个形状。

对 **SetShapesDefaultProperties** 对象使用 **Shape** 方法可设置文档的默认形状的格式。新形状将继承默认形状的许多属性。

使用 **Type** 属性可指定形状的类型；例如，任意多边形、自选图形、OLE 对象、标注或链接图片。使用 **AutoShapeType** 属性可指定自选图形的类型；例如，椭圆、矩形或气球。

使用 **Width** 和 **Height** 属性可指定形状的大小。

**TextFrame** 属性返回 **TextFrame** 对象，该对象包含用于将文本附加到形状和链接文本框之间的文本的所有属性和方法。

**Shape** 对象锁定到某一文本范围，但可以自由浮动，并且可以放置在页面上的任何位置。**InlineShape** 对象被视为字符，并作为字符置于文本行中。可以使用 **ConvertToInlineShape** 方法和 **ConvertToShape** 方法来转换形状的类型。只能将图片、OLE 对象和 ActiveX 控件转换为内嵌形状。

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**                      | 应用于使用 **PickUp** 方法复制的特定图形格式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropBottom**           | 从绘图画布的底部裁剪一定百分比的绘图画布高度。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropLeft**             | 从绘图画布左侧裁剪一定百分比的绘图画布宽度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropRight**            | 从绘图画布右侧裁剪一定百分比的绘图画布宽度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropTop**              | 从绘图画布顶部裁剪一定百分比的绘图画布高度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToInlineShape**       | 将文档绘图层的指定图形转换为文字层的嵌入式图形。只能转换代表图片、OLE 对象或 ActiveX 控件的图形。此方法返回一个 **InlineShape** 对象，该对象代表图片或 OLE 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                     | 删除指定的图形节点。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**                  | 创建指定的 **Shape** 对象的副本，以标准的偏移将新图形从原图形添加至 **Shapes** 集合，然后返回新的 **Shape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Flip**                       | 水平或垂直翻转一个图形。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementLeft**              | 将指定形状水平移动指定的磅数。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementRotation**          | 使指定的形状绕 Z 轴旋转指定的角度。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementTop**               | 以指定磅数垂直移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickUp**                     | 复制指定形状的格式。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleHeight**                | 以指定的比例缩放形状的高度。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleWidth**                 | 按指定的比例缩放形状的宽度。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                     | 选择指定的形状。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetShapesDefaultProperties** | 将文档中默认形状的格式应用于指定的形状。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup**                    | 取消组合指定形状中的所有组合形状。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**                     | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 Z 顺序中的位置）。 |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Adjustments**                | 返回一个 **Adjustments** 对象，该对象包含所有对指定 **Shape** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlternativeText**            | 返回或设置与网页的图形相关联的可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。）。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Anchor**                     | 返回一个 **Range** 对象，该对象代表指定图形或图形区域的锁定范围。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShapeType**              | 返回或设置指定的 **Shape** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。**MsoAutoShapeType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**            | 设置或返回指定形状的背景样式。可读/写 MsoBackgroundStyleIndex。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Callout**                    | 返回 **CalloutFormat** 对象，该对象包含指定图形的标注格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanvasItems**                | 返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**                      | 返回一个 **Chart** 对象，该对象代表文档中形状集合内的图表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Child**                      | 如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                       | 返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                       | 返回一个 **GlowFormat** 对象，该对象代表形状的发光格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**                 | 返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChart**                   | 如果指定的形状具有图表，则为 **True**。只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasSmartArt**                | 如果形状上显示 SmartArt 图表则返回 **True**。只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                     | 返回或设置指定图形的高度。**Single** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightRelative**             | 返回或设置一个 **Single** 类型的值，该值代表形状高度的相对百分比。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalFlip**             | 表示形状进行过水平翻转。**MsoTriState** 类型，只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlink**                  | 返回一个 **Hyperlink** 对象，该对象代表与 **Shape** 对象相关联的超链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                         | 返回指定形状的标识类型。只读 **Long** 类型。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutInCell**               | 返回一个 **Long** 类型的值，该值表示表格中的形状显示在表格内部还是表格外部。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                       | 返回或设置一个 **Single** 类型的值，该值代表指定形状或形状范围的水平位置（以磅为单位）。也可以是任何有效的 **WdShapePosition**常量。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftRelative**               | 返回或设置一个 **Single** 类型的值，该值代表形状左侧的相对位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**                       | 返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkFormat**                 | 返回一个 **LinkFormat** 对象，该对象代表链接到文件的形状的链接选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAnchor**                 | 如果 **Shape** 对象的锁定标记锁定到锁定范围，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAspectRatio**            | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                       | 返回或设置指定对象的名称。**String** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nodes**                      | 返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEFormat**                  | 返回一个 **OLEFormat** 对象，该对象代表指定形状、内嵌形状或域的 OLE 特性（链接除外）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回一个 **Object** 类型值，该值代表指定 **Shape** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentGroup**                | 返回一个 **Shape** 对象，该对象代表子形状或子形状范围的通用父形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**              | 返回一个 **PictureFormat** 对象，该对象包含指定对象的图片格式属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reflection**                 | 返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeHorizontalPosition** | 指定形状的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeHorizontalSize**     | 返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状区域相对的对象。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeVerticalPosition**   | 指定形状的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeVerticalSize**       | 返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状的相对垂直大小。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rotation**                   | 返回或设置指定图形绕 Z 轴旋转的度数。正数表示按顺时针方向进行旋转，负值表示按逆时针方向旋转。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Script**                     | 返回一个 **Script** 对象，该对象代表网页中图像的一段脚本或代码。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                     | 返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeStyle**                 | 返回或设置指定形状的形状样式。可读/写 MsoShapeStyleIndex。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartArt**                   | 返回 SmartArt 对象，该对象提供一种处理与指定形状相关联的 SmartArt 的方式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoftEdge**                   | 返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**                 | 返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame**                  | 返回一个 **TextFrame** 对象，该对象包含指定形状的文字。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**                 | 返回一个 **TextFrame2** 对象，该对象包含指定形状的文本。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**                     | 返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**                      | 返回或设置包含指定形状的标题的 **String** 类型值。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                        | 返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopRelative**                | 返回或设置一个 **Single** 类型的值，该值代表形状顶部的相对位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                       | 返回内嵌形状的类型。只读 **MsoShapeType** 类型。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalFlip**               | 如果指定形状围绕垂直轴进行翻转，则该属性值为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Vertices**                   | 该属性以一系列坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。只读 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                    | 如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                      | 返回或设置指定形状的宽度（以磅为单位）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WidthRelative**              | 返回或设置一个代表形状的相对宽度的 **Single**。可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WrapFormat**                 | 返回一个 **WrapFormat** 对象，该对象包含在指定的形状四周文字环绕的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrderPosition**             | 返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。 |

**成员方法**

#### **Shape.Apply**

应用于使用 **PickUp** 方法复制的特定图形格式。

**语法**

**express.Apply()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果以前没有使用 **PickUp** 方法复制 **Shape** 对象的格式，则使用 **Apply** 方法将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制活动文档中第一个图形的格式，然后将其应用到同一文档的第二个图形。*/ function test() { ActiveDocument.Shapes.Item(1).PickUp() ActiveDocument.Shapes.Item(2).Apply() }` |

#### **Shape.CanvasCropBottom**

从绘图画布的底部裁剪一定百分比的绘图画布高度。

**语法**

**express.CanvasCropBottom(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从底部裁剪绘图画布高度的百分之十。输入 0.1 从底部裁剪绘图画布高度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，则本示例从活动文档中第一块绘图画布的底部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function CropCanvasBottom() {     let shpCanvas = ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropBottom(0.75) }` |

#### **Shape.CanvasCropLeft**

从绘图画布左侧裁剪一定百分比的绘图画布宽度。

**语法**

**express.CanvasCropLeft(Increment )**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从左侧裁剪绘图画布宽度的百分之十。输入 0.1 从左侧裁剪绘图画布宽度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的左侧裁剪绘图画布宽度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function CropCanvasLeft() {     let shpCanvas = ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropLeft(0.75) }` |

#### **Shape.CanvasCropRight**

从绘图画布右侧裁剪一定百分比的绘图画布宽度。

**语法**

**express.CanvasCropRight(Increment )**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从右侧裁剪绘图画布宽度的百分之十。输入 0.1 从右侧裁剪绘图画布宽度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从右侧裁剪活动文档中第一块绘图画布宽度的百分之二十五。如果不是这种情况，需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function CropCanvasRight() {     let shpCanvas = ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropRight(0.75) }` |

#### **Shape.CanvasCropTop**

从绘图画布顶部裁剪一定百分比的绘图画布高度。

**语法**

**express.CanvasCropTop(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从顶部裁剪绘图画布高度的百分之十。输入 0.1 从顶部裁剪绘图画布高度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的顶部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在活动文档中添加绘图画布。 */ function CropCanvasTop(){     let shpCanvas = ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropTop(0.75) }` |

#### **Shape.ConvertToInlineShape**

将文档绘图层的指定图形转换为文字层的嵌入式图形。只能转换代表图片、OLE 对象或 ActiveX 控件的图形。此方法返回一个 **InlineShape** 对象，该对象代表图片或 OLE 对象。

**语法**

**express.ConvertToInlineShape()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

支持附加文字的图形不能转换为嵌入式图形。对这样的图形请使用 **ConvertToFrame** 方法。

如果对包含多个图形的 **ShapeRange** 对象使用此方法，将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 MyDoc.doc 中的每张图片转换为嵌入式图形。*/ function test() { for(let i = 1; i <= Documents.Item("MyDoc.doc").Shapes.Count; i++){     if(Documents.Item("MyDoc.doc").Shapes.Item(i).Type == msoPicture){        Documents.Item("MyDoc.doc").Shapes.Item(i).ConvertToInlineShape()     } } }  ` |

#### **Shape.Delete**

删除指定的图形节点。

**语法**

**express.Delete(Index)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Index*  | 必选          | **Long**     | 要删除的图形节点的数目。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*删除文档形状集合中的形状一*/ Application.ActiveDocument.Shapes.Item(1).Delete() ` |

#### **Shape.Duplicate**

创建指定的 **Shape** 对象的副本，以标准的偏移将新图形从原图形添加至 **Shapes** 集合，然后返回新的 **Shape** 对象。

**语法**

**express.Duplicate()**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建活动文档第一个图形的副本，然后改变新图形的填充格式。*/ function test() { let newShape = ActiveDocument.Shapes.Item(1).Duplicate()     newShape.Fill.PresetGradient(msoGradientVertical, 1, msoGradientGold) }` |

#### **Shape.Flip**

水平或垂直翻转一个图形。

**语法**

**express.Flip(FlipCmd)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**   | **说明**   |
| --------- | ------------- | -------------- | ---------- |
| *FlipCmd* | 必选          | **MsoFlipCmd** | 翻转方向。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先向活动文档添加一个三角形，然后复制此三角形，再垂直翻转复制的三角形并将其设为红色。*/ function FlipShape(){     let aShape = ActiveDocument.Shapes.AddShape(msoShapeRightTriangle, 150, 150, 50, 50).Duplicate()     aShape.Fill.ForeColor.RGB = (255, 0, 0)     aShape.Fill(msoFlipVertical) }` |

#### **Shape.IncrementLeft**

将指定形状水平移动指定的磅数。

**语法**

**express.IncrementLeft(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状水平移动的距离，以磅为单位。为正值时将形状右移；为负值时将形状左移。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/ function test() { let myDocument = ActiveDocument     myDocument.Shapes.Item(1).Duplicate     myDocument.Fill.PresetTextured(msoTextureGranite)     myDocument.IncrementLeft(70)     myDocument.IncrementTop(-50)     myDocument.IncrementRotation(30) }` |

#### **Shape.IncrementRotation**

使指定的形状绕 Z 轴旋转指定的角度。

**语法**

**express.IncrementRotation(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状的水平旋转量，以度为单位。为正值时顺时针旋转形状，为负值时逆时针旋转形状。 |

**说明**

使用 **Rotation** 属性可以设置形状的绝对旋转角度。要绕 X 轴或 Y 轴旋转三维形状，请使用 **ThreeDFormat** 对象的 **IncrementRotationX** 或 **IncrementRotationY** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/ function test() { let myDocument = ActiveDocument     myDocument.Shapes.Item(1).Duplicate     myDocument.Fill.PresetTextured(msoTextureGranite)     myDocument.IncrementLeft(70)     myDocument.IncrementTop(-50)     myDocument.IncrementRotation(30) }  ` |

#### **Shape.IncrementTop**

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop(Increment)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定对象的垂直移动距离，以磅为单位。为正值时将形状下移；为负值时将形状上移。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/ function test() { let myDocument = ActiveDocument     myDocument.Shapes.Item(1).Duplicate     myDocument.Fill.PresetTextured(msoTextureGranite)     myDocument.IncrementLeft(70)     myDocument.IncrementTop(-50)     myDocument.IncrementRotation(30) }  ` |

#### **Shape.PickUp**

复制指定形状的格式。

**语法**

**express.PickUp()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

用 **Apply** 方法可将复制的格式应用于另一个形状。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例首先复制 myDocument 上第一个形状的格式，然后将复制的形状格式应用于第二个形状。*/ function test() { let myDocument = ActiveDocument     myDocument.Shapes.Item(1).PickUp()     myDocument.Shapes.Item(2).Apply() }` |

#### **Shape.ScaleHeight**

以指定的比例缩放形状的高度。

**语法**

**express.ScaleHeight(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的高度与当前或原始高度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。对于图片和 OLE 对象以外的形状总是相对于当前高度缩放。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 上的所有图片和 OLE 对象放大至原始高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/ function test() { let myDocument = ActiveDocument let s for(let i = 1; i <= myDocument.Shapes.Count;i++){     s = myDocument.Shapes.Item(i)     switch(s.Type){         case msoEmbeddedOLEObject:         case msoLinkedOLEObject:         case msoOLEControlObject:         case msoLinkedPicture:         case msoPicture:             s.ScaleHeight(1.75, true)             s.ScaleWidth(1.75, true)             break         default:             s.ScaleHeight(1.75, false)             s.ScaleWidth(1.75, false)     } } }` |

#### **Shape.ScaleWidth**

按指定的比例缩放形状的宽度。

**语法**

**express.ScaleWidth(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。图片和 OLE 对象以外的形状总是相对于当前宽度缩放。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 上的所有图片和 OLE 对象放大至原始高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/ function test() { let myDocument = ActiveDocument let s for(let i = 1; i <= myDocument.Shapes.Count;i++){     s = myDocument.Shapes.Item(i)     switch(s.Type){         case msoEmbeddedOLEObject:         case msoLinkedOLEObject:         case msoOLEControlObject:         case msoLinkedPicture:         case msoPicture:             s.ScaleHeight(1.75, true)             s.ScaleWidth(1.75, true)             break         default:             s.ScaleHeight(1.75, false)             s.ScaleWidth(1.75, false)     } } }` |

#### **Shape.Select**

选择指定的形状。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | 在添加形状时，如果该参数值为 **True**，则替换所选内容。如果该参数值为 **False**，则将新形状添加到所选内容。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*选中文档形状集合中的形状一*/ Application.ActiveDocument.Shapes.Item(1).Select()` |

#### **Shape.SetShapesDefaultProperties**

将文档中默认形状的格式应用于指定的形状。

**语法**

**express.SetShapesDefaultProperties()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

新形状将继承默认形状的许多属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一个矩形添至 myDocument，设置矩形的填充格式，将矩形的格式应用于默认形状，然后将另一个（较小的）矩形添至文档。后一个矩形与前一个矩形有相同的填充格式。*/ function test() { let mydocument = ActiveDocument let Shape = mydocument.Shapes let newShape1 = Shape.AddShape(msoShapeRectangle, 5, 5, 80, 60)     let newShape2 = newShape1.Fill         newShape2.ForeColor.RGB = (0, 0, 255)         newShape2.BackColor.RGB = (0, 204, 255)         newShape2.Patterned(msoPatternHorizontalBrick) 	//Sets formatting for default shapes 	newShape1.SetShapesDefaultProperties() //New shape has default formatting Shape.AddShape(msoShapeRectangle, 90, 90, 40, 30) }` |

#### **Shape.Ungroup**

取消组合指定形状中的所有组合形状。

**语法**

**express.Ungroup()**

*express*   一个代表 **Shape** 对象的变量。

**返回值**

ShapeRange

**说明**

此方法分解指定形状中的图片和 OLE 对象并将取消组合后的形状作为单个 **ShapeRange** 对象返回。

因为将组合形状作为单个对象处理，所以对形状进行组合或取消组合操作时，将会更改 **Shapes** 集合中项的数目，并更改集合中受影响的项之后的项的索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例取消组合 myDocument 中的所有组合形状，并分解所有的图片和 OLE 对象。*/ function test() { let myDocument = ActiveDocument for(let i = 1; i <= myDocument.Shapes.Count; i++){     myDocument.Shapes.Item(i).Ungroup() } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将取消组合 myDocument 中的所有组合形状，但不分解文档中图片或 OLE 对象。*/ function test() { let myDocument = ActiveDocument for(let i = 1; i <= myDocument.Shapes.Count; i++){     if(myDocument.Shapes.Item(i).Type == msoGroup){        myDocument.Shapes.Item(i).Ungroup()     } } }` |

#### **Shape.ZOrder**

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 Z 顺序中的位置）。

**语法**

**express.ZOrder(ZOrderCmd)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                   |
| ----------- | ------------- | ---------------- | ------------------------------------------ |
| *ZOrderCmd* | 必选          | **MsoZOrderCmd** | 指定相对于其他形状将指定形状移动到的位置。 |

**返回值**

无

**说明**

使用 **ZOrderPosition** 属性可确定形状在 Z 顺序中的当前位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中添加一个椭圆形，并将该椭圆形置于 Z 顺序中倒数第二的位置（如果文档中至少还有另一个图形）。*/ function test() { let zorderposition = ActiveDocument.Shapes.AddShape(msoShapeOval,  100, 100, 100, 300) while(zorderposition.ZOrderPosition >  2){       zorderposition.ZOrder(msoSendBackward) } }` |

**成员属性**

#### **Shape.Adjustments**

返回一个 **Adjustments** 对象，该对象包含所有对指定 **Shape** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。

**语法**

**express.Adjustments**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上图形 3 的调整值 1 设为 0.25。*/ ActiveDocument.Shapes.Item(3).Adjustments.Item(1, 0.25)` |

#### **Shape.AlternativeText**

返回或设置与网页的图形相关联的可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。）。**String** 类型，可读写。

**语法**

**express.AlternativeText**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Anchor**

返回一个 **Range** 对象，该对象代表指定图形或图形区域的锁定范围。只读。

**语法**

**express.Anchor**

*express*   一个代表 **Shape** 对象的变量。

**说明**

所有 **Shape** 对象都锁定于某一个文字区域，但可将其置于锁定标记所在页的任意位置。如果创建图形时指定了锁定范围，则锁定标记位于包含该锁定范围的第一段落的段首。如果没有指定锁定范围，则将自动选择锁定范围，图形参照页面的左边框和上边框进行定位。

图形总是与锁定标记处在同一页上。如果图形的 **LockAnchor** 属性为 **True**，则不能在页面上拖动锁定标记。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定活动文档第一个图形锁定标记所在的段落。*/ ActiveDocument.Shapes.Item(1).Anchor.Paragraphs.Item(1).Range.Select()` |

#### **Shape.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Shape** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Shape.AutoShapeType**

返回或设置指定的 **Shape** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。**MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

*express*   一个代表 **Shape** 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的所有 16 磅星形替换为 32 磅星形。*/ function ReplaceAutoShape(){     let shpStar     let docNew = ActiveDocument     for(let i = 1; i <= docNew.Shapes.Count; i++){         shpStar = docNew.Shapes.Item(i)         if(shpStar.AutoShapeType == msoShape16pointStar){         shpStar.AutoShapeType = msoShape32pointStar         }     } }` |

#### **Shape.BackgroundStyle**

设置或返回指定形状的背景样式。可读/写 MsoBackgroundStyleIndex。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Callout**

返回 **CalloutFormat** 对象，该对象包含指定图形的标注格式属性。只读。

**语法**

**express.Callout**

*express*   一个代表 **Shape** 对象的变量。

**说明**

该属性应用于代表标注的 **Shape** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例在 myDocument 中添加一个椭圆和一个指向椭圆的标注。标注文本不带有边框，但有一条强调线将文本和标注线隔开。*/ function test() { let Mydocument = ActiveDocument.Shapes     Mydocument.AddShape(msoShapeOval, 180, 200, 280, 130) let myDocument = Mydocument.AddCallout(msoCalloutTwo, 420, 170, 170, 40)     myDocument.TextFrame.TextRange.Text = "My oval" let newmyDocument = myDocument.Callout     newmyDocument.Accent = true     newmyDocument.Border = false }` |

#### **Shape.CanvasItems**

返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。

**语法**

**express.CanvasItems**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中创建一个新的绘图画布，并在其上添加一个圆圈。*/ function NewCanvasShape(){     let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200) 	shpCanvas.CanvasItems.AddShape(msoShapeOval, 25, 25, 150, 150) }` |

#### **Shape.Chart**

返回一个 **Chart** 对象，该对象代表文档中形状集合内的图表。只读。

**语法**

**express.Chart**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Child**

如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.Child**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择绘图画布中的第一个图形，并且如果所选图形是一个子图形，用指定颜色填充图形。本示例假定活动文档中的第一个图形是包含多个图形的绘图画布。*/ function FillChildShape(){     //Select the first shape in the drawing canvas     let shpCanvasItem = ActiveDocument.Shapes.Item(1).CanvasItems.Item(1)      //Fill selected shape if it is a child shape     if(shpCanvasItem.Child == msoTrue){         shpCanvasItem.Fill.ForeColor.RGB = 100, 0, 200     }     else {         MsgBox("This shape is not a child shape.")     } }` |

#### **Shape.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Shape.Fill**

返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先向 myDocument 添加一个矩形，然后为其设置前景色、背景色和矩形填充的渐变类型。*/ function test() { let myDocument = Documents.Item(1) myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill myDocument.ForeColor.RGB = 128, 0, 0 myDocument.BackColor.RGB = 170, 170, 170 myDocument.TwoColorGradient(msoGradientHorizontal, 1) }` |

#### **Shape.Glow**

返回一个 **GlowFormat** 对象，该对象代表形状的发光格式。只读。

**语法**

**express.Glow**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.GroupItems**

返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **Shape** 对象的变量。

**说明**

该属性应用于代表组合图形的 **Shape** 对象。使用 **GroupShapes** 对象的 **Item** 方法可从图形组中返回单个图形。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加三个三角形，并组合这些三角形，为整个图形组设置颜色，然后又独立地更改第二个三角形的颜色。*/ function test() { let myShapes = ActiveDocument.Shapes  myShapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne" myShapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo" myShapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"  let gruopShapes = myShapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()  gruopShapes.Fill.PresetTextured(msoTextureBlueTissuePaper) gruopShapes.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)  }` |

#### **Shape.HasChart**

如果指定的形状具有图表，则为 **True**。只读。

**语法**

**express.HasChart**

*express*   一个代表 **Shape** 对象的变量。

**说明**

对于 OLE 图表，此属性总是返回 False。此时请使用 `InlineShape.OLEFormat.ProgID`，并且检查是否存在下列可能的值：“Excel.Chart.8”、“MSGraph.Chart.8”、“Excel.Sheet.8”、“Excel.Chart.5”、“MSGraph.Chart.5”或“Excel.Sheet.5”。

#### **Shape.HasSmartArt**

如果形状上显示 SmartArt 图表则返回 **True**。只读。

**语法**

**express.HasSmartArt**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中的第一个形状是否包含 SmartArt。*/ function test() { let myShape = ActiveDocument.Shapes.Item(1) if(myShape.HasSmartArt){     MsgBox("The first shape contains SmartArt.") } else{     MsgBox("The first shape contains no SmartArt.") } }` |

#### **Shape.Height**

返回或设置指定图形的高度。**Single** 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一张图片作为嵌入式图形插入活动文档，并更改此图片的高度和宽度*/ function test() {     let aInLine = Application.ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Bubbles.bmp", Selection.Range)     aInLine.Height = 100     aInLine.Width = 200 }` |

#### **Shape.HeightRelative**

返回或设置一个 **Single** 类型的值，该值代表形状高度的相对百分比。可读写。

**语法**

**express.HeightRelative**

*express*   一个代表 **Shape** 对象的变量。

**说明**

将此属性与**RelativeVerticalSize**属性一起使用。当设置为 **wdShapeSizeRelativeNone** (-999999)（参见**WdShapeSizeRelative**枚举）时，应忽略此属性，因为形状不能使用百分比大小。高度是由 **Height** 属性单独确定的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*设置形状一的HeightRelative属性为50*/ Application.ActiveDocument.Shapes.Item(1).HeightRelative = 50` |

#### **Shape.HorizontalFlip**

表示形状进行过水平翻转。**MsoTriState** 类型，只读。

**语法**

**express.HorizontalFlip**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/ function FlipShape(){     for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){         if(ActiveDocument.Shapes.Item(i).HorizontalFlip){             ActiveDocument.Shapes.Item(i).Flip(msoFlipHorizontal)         }         if(ActiveDocument.Shapes.Item(i).VerticalFlip){             ActiveDocument.Shapes.Item(i).Flip(msoFlipVertical)         }     } }` |

#### **Shape.Hyperlink**

返回一个 **Hyperlink** 对象，该对象代表与 **Shape** 对象相关联的超链接。只读。

**语法**

**express.Hyperlink**

*express*   一个代表 **Shape** 对象的变量。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果不存在与指定形状相关联的超链接，则会发生错误。在此情况下，请将 Add 方法用于 Hyperlinks 集合来为指定形状添加超链接。以下示例说明了具体的操作方法。*/ ActiveDocument.Hyperlinks.Add(Selection.Shapes.Item(1), "http://www.microsoft.com")` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中第一个形状的超链接地址。*/ MsgBox(ActiveDocument.Shapes.Item(1).Hyperlink.Address)` |

#### **Shape.ID**

返回指定形状的标识类型。只读 **Long** 类型。

**语法**

**express.ID**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*获取形状的ID*/ Application.ActiveDocument.Shapes.Item(1).ID` |

#### **Shape.LayoutInCell**

返回一个 **Long** 类型的值，该值表示表格中的形状显示在表格内部还是表格外部。

**语法**

**express.LayoutInCell**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果该属性为 **True**，则表示指定的图片显示在表格内部；如果该属性为 **False**，则表示指定的图片显示在表格外部。

**LayoutInCell** 属性对应于图片格式的**“高级版式”**对话框中的**“表格单元格中的版式”**选项。

| 注释                                                         |
| ------------------------------------------------------------ |
| 仅当 **WrapFormat** 对象的 **Type** 属性设置为除 **wdWrapTypeInline** 或 **wdWrapTypeNone** 之外的值时，对 **LayoutInCell** 属性的设置才会生效。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对活动文档中的第一个形状禁用“表格单元格中的版式”选项。本示例假设指定的形状在表格内并且不是嵌入形状。*/ ActiveDocument.Shapes.Item(1).LayoutInCell = false` |

#### **Shape.Left**

返回或设置一个 **Single** 类型的值，该值代表指定形状或形状范围的水平位置（以磅为单位）。也可以是任何有效的 **WdShapePosition**常量。可读写。

**语法**

**express.Left**

*express*   一个代表 **Shape** 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。**RelativeHorizontalPosition** 属性控制锁定标记是沿字符、文本栏、页边距还是页面边缘放置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个图形的水平位置设置为距页面的左边缘 1 英寸*/  function test() {     let shapes = Application.ActiveDocument.Shapes.Item(1)     shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage     shapes.Left = InchesToPoints(1) }` |

#### **Shape.LeftRelative**

返回或设置一个 **Single** 类型的值，该值代表形状左侧的相对位置。可读写。

**语法**

**express.LeftRelative**

*express*   一个代表 **Shape** 对象的变量。

**说明**

将此属性与 **RelativeHorizontalPosition** 属性一起使用。当设置为 **wdShapePositionRelativeNone** (-999999)（参见 **WdShapePositionRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比放置。水平位置是由 **Left** 属性单独确定的。

#### **Shape.Line**

返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。

**语法**

**express.Line**

*express*   一个代表 **Shape** 对象的变量。

**说明**

对于线条来说，**LineFormat** 对象代表线条本身；而对于带有边框的形状来说，**LineFormat** 对象代表边框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 添加一条蓝色虚线。*/ function test() { let line = ActiveDocument.Shapes.AddLine(10, 10, 250, 250).Line line.DashStyle = msoLineDashDotDot line.ForeColor.RGB = (50, 0, 128) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 添加一个十字形，然后将其边框粗细设置为 8 磅、颜色为红色。*/ function test() { let line = ActiveDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line line.Weight = 8 line.ForeColor.RGB = (255, 0, 0) }` |

#### **Shape.LinkFormat**

返回一个 **LinkFormat** 对象，该对象代表链接到文件的形状的链接选项。只读。

**语法**

**express.LinkFormat**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一个图形作为内嵌形状插入（使用 INCLUDEPICTURE 域），然后显示其源名称 (Tiles.bmp)。*/ function test() { let iShape = ActiveDocument.InlineShapes.AddPicture("C:\\windows\\Tiles.bmp", true, false, Selection.Range) MsgBox(iShape.LinkFormat.SourceName) }  ` |

#### **Shape.LockAnchor**

如果 **Shape** 对象的锁定标记锁定到锁定范围，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.LockAnchor**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果形状有一个锁定的标记，则不能通过拖动来移动形状的锁定标记。锁定标记不会随形状移动而移动。虽然 **Shape** 对象已锁定到某一文字范围，但是您可以将该对象放在页面上的任何位置。形状锁定到包含锁定范围的第一段的开始。形状总是与其锁定标记位于同一页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一个文档，为其添加一个形状，然后锁定形状的锁定标记。*/ function test() { let myDoc = Documents.Add() let myShape = myDoc.Shapes.AddShape(msoShapeBalloon, 100, 100, 140, 70) myShape.LockAnchor = true ActiveDocument.ActiveWindow.View.ShowObjectAnchors = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回一条消息，报告活动文档中每个形状的锁定状态。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     MsgBox("Shape "+ i + " is locked - " +ActiveDocument.Shapes.Item(i).LockAnchor) } }` |

#### **Shape.LockAspectRatio**

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 添加一个立方体。可以移动此立方体并调整其大小，但不能重新设置其比例。*/ function test() { let myDocument = ActiveDocument myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200).LockAspectRatio = msoTrue }` |

#### **Shape.Name**

返回或设置指定对象的名称。**String** 类型，可读写。

**语法**

**express.Name**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Nodes**

返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。

**语法**

**express.Nodes**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档的图形 3 的第四个顶点的后面添加一个带有一条曲线的平滑顶点。图形 3 必须为至少带有四个顶点的任意多边形。*/ function test() { let nodes = ActiveDocument.Shapes.Item(3).Nodes nodes.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100) }` |

#### **Shape.OLEFormat**

返回一个 **OLEFormat** 对象，该对象代表指定形状、内嵌形状或域的 OLE 特性（链接除外）。只读。

**语法**

**express.OLEFormat**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例遍历活动文档中的所有浮动形状，并将所有链接的 ET 工作表设置为自动更新。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){         if(ActiveDocument.Shapes.Item(i).OLEFormat.ProgID == "Excel.Sheet"){             ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = ture         }     } } }` |

#### **Shape.Parent**

返回一个 **Object** 类型值，该值代表指定 **Shape** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.ParentGroup**

返回一个 **Shape** 对象，该对象代表子形状或子形状范围的通用父形状。

**语法**

**express.ParentGroup**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中创建两个形状并对其进行分组。然后利用一个形状访问该组的父组，并使用同样的填充颜色填充父组中的所有形状。本示例假定活动文档当前并不包括形状，如果包括形状，则产生错误。*/ function ParentGroupShape(){     //Add two shapes to active document and group     let shapes = ActiveDocument.Shapes     shapes.AddShape(msoShapeOval, 72, 72, 100, 100)     shapes.AddShape(msoShapeHeart, 110, 120, 100, 100)     shapes.Range(Array(1, 2)).Group()     let pgShape = ActiveDocument.Shapes.Item(1).GroupItems.Item(1).ParentGroup     pgShape.Fill.ForeColor.RGB = (100, 0, 255) }` |

#### **Shape.PictureFormat**

返回一个 **PictureFormat** 对象，该对象包含指定对象的图片格式属性。

**语法**

**express.PictureFormat**

*express*   一个代表 **Shape** 对象的变量。

**说明**

该属性适用于代表图片或 OLE 对象的 **Shape** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/ function test() { let myDocument = ActiveDocument.Shapes.Item(1).PictureFormat myDocument.Brightness = 0.3 myDocument.Contrast = .75 }` |

#### **Shape.Reflection**

返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.RelativeHorizontalPosition**

指定形状的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。

**语法**

**express.RelativeHorizontalPosition**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.RelativeHorizontalSize**

返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeHorizontalSize**

*express*   一个代表 **Shape** 对象的变量。

**说明**

此属性可与 **WidthRelative** 属性一起使用。

#### **Shape.RelativeVerticalPosition**

指定形状的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。

**语法**

**express.RelativeVerticalPosition**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例重新定位活动文档中第一个形状对象。*/ function test() { let shapes = ActiveDocument.Shapes.Item(1) shapes.Left = InchesToPoints(0.6) shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage shapes.Top = InchesToPoints(1) shapes.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph }` |

#### **Shape.RelativeVerticalSize**

返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状的相对垂直大小。可读写。

**语法**

**express.RelativeVerticalSize**

*express*   一个代表 **Shape** 对象的变量。

**说明**

此属性可与 **HeightRelative** 属性一起使用。

#### **Shape.Rotation**

返回或设置指定图形绕 Z 轴旋转的度数。正数表示按顺时针方向进行旋转，负值表示按逆时针方向旋转。**Single** 类型，可读写。

**语法**

**express.Rotation**

*express*   一个代表 **Shape** 对象的变量。

**说明**

要设置三维图形绕 X 轴或 Y 轴的旋转，请使用 **ThreeDFormat** 对象的 **RotationX** 属性或 **RotationY** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 中所有形状的旋转量设置为与形状 1 的旋转量一致。*/ function test() { let aShape = ActiveDocument.Shapes let sh1Rotation = aShape.Item(1).Rotation for(let i = 1; i <= aShape.Count; i++){     aShape.Item(i).Rotation = sh1Rotation } }  ` |

#### **Shape.Script**

返回一个 **Script** 对象，该对象代表网页中图像的一段脚本或代码。

**语法**

**express.Script**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果该网页中不包含脚本，则不返回任何内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中第一个形状所用的脚本语言的类型。*/ function test() { let objScr = ActiveDocument.Shapes.Item(1).script if(objScr){     switch(objScr.Language){         case msoScriptLanguageVisualBasic:             MsgBox("VBScript")             break         case msoScriptLanguageVisualBasic:             MsgBox("JavaScript")             break         case msoScriptLanguageVisualBasic:             MsgBox("Active Server Pages")             break         default:             Msgbox("Other scripting language")     } } }  ` |

#### **Shape.Shadow**

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。

**语法**

**express.Shadow**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向活动文档中添加一个带阴影格式的箭头。*/ function test() { let myShape = ActiveDocument.Shapes.AddShape(msoShapeRightArrow, 90, 79, 64, 43) myShape.Shadow.Type = msoShadow5 }` |

#### **Shape.ShapeStyle**

返回或设置指定形状的形状样式。可读/写 MsoShapeStyleIndex。

**语法**

**express.ShapeStyle**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例更改活动文档中第一个形状的形状样式。*/ function test() { let myShape = ActiveDocument.Shapes.Item(1) myShape.ShapeStyle = msoLineStylePreset12 }` |

#### **Shape.SmartArt**

返回 SmartArt 对象，该对象提供一种处理与指定形状相关联的 SmartArt 的方式。只读。

**语法**

**express.SmartArt**

*express*   一个代表 **Shape** 对象的变量。

**说明**

**SmartArt** 属性为与形状相关联的 SmartArt 图形进行交互提供了一个入口点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中添加 SmartArt 图形。*/ function test() { let myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 100, 100, 400, 400) let mySmartArt = myShape.SmartArt }` |

#### **Shape.SoftEdge**

返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。

**语法**

**express.SoftEdge**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.TextEffect**

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

*express*   一个代表 **Shape** 对象的变量。

**说明**

该属性适用于代表艺术字的 **Shape** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，如果 myDocument 的第三个形状为艺术字，则将它的字形设为加粗。*/ function test() { let myShape = ActiveDocument.Shapes.Item(3) if(myShape.Type == msoTextEffect){     myShape.TextEffect.FontBold = true } }  ` |

#### **Shape.TextFrame**

返回一个 **TextFrame** 对象，该对象包含指定形状的文字。

**语法**

**express.TextFrame**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在 myDocument 中添加一个矩形，在该矩形中添加文本，并为文本框架设置边距。*/ function test() { let textframe = ActiveDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame textframe.TextRange.Text = "Here is some test text" textframe.MarginBottom = 0 textframe.MarginLeft = 100 textframe.MarginRight = 0 textframe.MarginTop = 20 }  ` |

#### **Shape.TextFrame2**

返回一个 **TextFrame2** 对象，该对象包含指定形状的文本。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.ThreeD**

返回一个 **ThreeDFormat** 对象，该对象包含指定形状的三维格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置应用于 myDocument 中第一个形状的三维效果的深度、延伸颜色、延伸方向以及照明方向。*/ function test() { let myDocument = ActiveDocument.Shapes.Item(1).ThreeD myDocument.Visible = true myDocument.Depth = 50 // RGB value for purple myDocument.ExtrusionColor.RGB = (255, 100, 255) myDocument.SetExtrusionDirection(msoExtrusionTop) myDocument.PresetLightingDirection = msoLightingLeft }  ` |

#### **Shape.Title**

返回或设置包含指定形状的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

*express*   一个代表 **Shape** 对象的变量。

**说明**

使用 **Title** 属性可为形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中**“设置形状格式”**对话框的**“可选文字”**窗格上的**“标题”**文本框中。

| 注释                                                         |
| ------------------------------------------------------------ |
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中的第二个形状中添加可选文字标题。*/ ActiveDocument.Shapes.Item(2).Title = "Shape 2."` |

#### **Shape.Top**

返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Top**

*express*   一个代表 **Shape** 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。**RelativeVerticalPosition** 属性控制形状的锁定标记是沿行、段落、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象，**Top** 属性设置每个形状的垂直位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个形状的垂直位置设置为距页面顶部 1 英寸*/  function test() {     let shapes = Application.ActiveDocument.Shapes.Item(1)     shapes.RelativeVerticalPosition = wdRelativeVerticalPositionPage     shapes.Top = InchesToPoints(1) }` |

#### **Shape.TopRelative**

返回或设置一个 **Single** 类型的值，该值代表形状顶部的相对位置。可读写。

**语法**

**express.TopRelative**

*express*   一个代表 **Shape** 对象的变量。

**说明**

将此属性与 **RelativeHorizontalPosition** 属性一起使用。当设置为 **wdShapePositionRelativeNone** (-999999)（参见 **WdShapePositionRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比放置。垂直位置是由 **Top** 属性单独确定的。

#### **Shape.Type**

返回内嵌形状的类型。只读 **MsoShapeType** 类型。

**语法**

**express.Type**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.VerticalFlip**

如果指定形状围绕垂直轴进行翻转，则该属性值为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).HorizontalFlip == -1){         ActiveDocument.Shapes.Item(i).Flip(msoFlipHorizontal)     }     if(ActiveDocument.Shapes.Item(i).VerticalFlip == -1){         ActiveDocument.Shapes.Item(i).Flip(msoFlipVertical)     } } }  ` |

#### **Shape.Vertices**

该属性以一系列坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。只读 **Variant** 类型。

**语法**

**express.Vertices**

*express*   一个代表 **Shape** 对象的变量。

**说明**

可将该属性返回的数组用作 **AddCurve** 或 **AddPolyLine** 方法的参数。

下表显示 **Vertices** 属性如何将数组 *vertArray()* 的值与三角形的顶点坐标相关联。

vertArray

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*第一个顶点至文档的左边界的水平距离。*/ vertArray(1, 1)   /*第一个顶点至文档顶端的垂直距离。*/ vertArray(1, 2)   /*第二个顶点至文档的左边界的水平距离。*/ vertArray(2, 1)   /*第二个顶点至文档顶端的垂直距离。*/ vertArray(2, 2)   /*第三个顶点至文档的左边界的水平距离。*/ vertArray(3, 1)   /*第三个顶点至文档顶端的垂直距离。*/ vertArray(3, 2) ` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个形状的顶点坐标赋给一个数组变量，并显示第一个顶点的坐标。第一个形状必须是任意多边形图形。*/ function test() { let ishape = ActiveDocument.Shapes.Item(1) vertArray = ishape.Vertices x1 = vertArray[1][1] y1 = vertArray[1][2] MsgBox("First vertex coordinates: "+ x1 + "," + y1) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例创建一条与活动文档中第一个形状具有相同几何特征的曲线。该示例假定第一个形状为包含 3n+1 个顶点的贝赛尔曲线，其中 n 为曲线段数量。*/ function test() { let shapes = ActiveDocument.Shapes shapes.AddCurve(shapes.Item(1).Vertices, Selection.Range) }` |

#### **Shape.Visible**

如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True**。**MsoTriState** 类型，可读写。

**语法**

**express.Visible**

*express*   一个代表 **Shape** 对象的变量。

**说明**

如果 **Visible** 属性为 **False**，则某些方法和属性可能无法使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一篇文档，然后在文档中添加文本和一个矩形，同时将 WPS 设置为在打印文档时隐藏该矩形，然后在打印结束后再显示该图形。*/ function test() { let myDoc = Documents.Add(Selection.TypeText("This is some sample text.")) myDoc.Shapes.AddShape(msoShapeRectangle, 200, 70, 150, 60) myDoc.Shapes.Item(1).Visible = false myDoc.PrintOut() myDoc.Shapes.Item(1).Visible = true }` |

#### **Shape.Width**

返回或设置指定形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.WidthRelative**

返回或设置一个代表形状的相对宽度的 **Single**。可读写。

**语法**

**express.WidthRelative**

*express*   一个代表 **Shape** 对象的变量。

**说明**

将此属性与 **RelativeVerticalSize** 属性一起使用。当设置为 **wdShapeSizeRelativeNone** (-999999)（参见 **WdShapeSizeRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比大小。宽度是由 **Width** 属性单独确定的。

#### **Shape.WrapFormat**

返回一个 **WrapFormat** 对象，该对象包含在指定的形状四周文字环绕的属性。只读。

**语法**

**express.WrapFormat**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中添加一个椭圆，并指定文档文字环绕该椭圆外切矩形的左右两侧。该示例将文档文字与该矩形四边之间的边距设置为 0.1 英寸。*/ function test() { let myOval = ActiveDocument.Shapes.AddShape(msoShapeOval, 36, 36, 90, 50).WrapFormat myOval.Type = wdWrapSquare myOval.Side = wdWrapBoth myOval.DistanceTop = InchesToPoints(0.1) myOval.DistanceBottom = InchesToPoints(0.1) myOval.DistanceLeft = InchesToPoints(0.1) myOval.DistanceRight = InchesToPoints(0.1) }` |

#### **Shape.ZOrderPosition**

返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。

**语法**

**express.ZOrderPosition**

*express*   一个代表 **Shape** 对象的变量。

**说明**

`Shapes(1)` 返回 Z 顺序中的最后一个形状，而 `Shapes(Shapes.Count)` 返回 Z 顺序中的第一个形状。该属性为只读。要设置形状在 Z 顺序中的位置，请使用 **ZOrder** 方法。

形状在 Z 顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果在 myDocument 上有四个形状，则表达式 `myDocument.Shapes(1)` 返回 Z 顺序中的最后一个形状，而表达式 `myDocument.Shapes(4)` 返回 Z 顺序中的第一个形状。

无论何时将新形状添加到集合中，默认情况下该形状都将添加到 Z 顺序的最前端。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例首先向 myDocument 中添加一个椭圆形。如果该文档另外至少还有一个形状，则按照 Z 顺序将此椭圆形放置在倒数第二的位置。*/ function test() { let myDocument = ActiveDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) while(myDocument.ZOrderPosition > 2){     myDocument.ZOrder(msoSendBackward) } }` |

适用环境：web

适用平台：windows/linux