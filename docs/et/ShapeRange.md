**ShapeRange**



代表形状区域，它是文档中的一组形状。

**说明**

形状区域可以只包含文档中的一个形状，或者也可包含所有形状。您可以在形状区域中包含所需的任意形状（在文档中的所有形状中选取，或从选定内容中的所有形状中选取）。例如，您可以构造一个 **ShapeRange** 集合，它包含文档中前三个形状、文档中所有选定的形状，或文档中所有的任意多边形的。

**1. 返回指定名称或索引号的一组形状**

使用 `Shapes.Range`(*index*)（其中 *index* 是形状的名称或索引号，或由形状的名称或索引号组成的数组）可返回代表文档中的一组形状的 **ShapeRange** 集合。您可以使用 **Array** 函数来构造名称或索引号的数组。下例设置 *myDocument* 上形状一和三的填充图案。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick) }` |

下例设置 *myDocument* 上名为“Oval 4”和“Rectangle 5”的形状的填充图案。

虽然可以使用 **Range** 属性来返回任意数量的形状或幻灯片，但如果您只想返回一个集合成员，则使用 **Item** 方法会更简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let myRange = myDocument.Shapes.Range(["Oval 4", "Rectangle 5"])     myRange.Fill.Patterned(msoPatternHorizontalBrick) }` |

**2. 返回文档中全部或部分选定的形状**

使用 **Selection** 对象的 **ShapeRange** 属性可返回选定对象中的所有形状。下例设置窗口一中选定对象的所有形状的填充前景色（假定所选内容中至少有一个形状）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Windows.Item(1).Selection.ShapeRange.Fill.ForeColor.RGB = (255, 0, 255)` |

使用 `Selection.ShapeRange`(*index*)（其中 *index* 是形状的名称或索引号）返回某一选定的形状。下例设置了窗口一内选定形状的集合中形状二的前景填充色（假定该选定内容中至少有两个形状）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Windows.Item(1).Selection.ShapeRange.Item(2).Fill.ForeColor.RGB = (255, 0, 255)` |

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Align**                      | 对齐指定形状区域中的形状。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**                      | 应用通过 **PickUp**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.PickUp)方法复制的指定形状格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                     | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Distribute**                 | 水平或垂直地分布指定的形状区域中的各形状。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**                  | 复制对象，并返回对新复制对象的引用。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Flip**                       | 绕指定形状的水平或垂直对称轴翻转该形状。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Group**                      | 将指定区域中的形状形成一组。返回值：一个代表组合形状的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementLeft**              | 以指定磅数水平移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementRotation**          | 使指定的形状按指定度数值绕 Z 轴旋转。使用 **Rotation**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Rotation)属性可设置形状的绝对转角。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementTop**               | 以指定磅数垂直移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                       | 从集合中返回一个对象。返回值：包含在集合中的一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickUp**                     | 复制指定形状的格式。使用 **Apply**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Apply)方法可将复制的格式应用到其他形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Regroup**                    | 将指定形状原先所属的形状区域重新组合。以单一的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象返回重新组合的形状。返回值类型：Shape |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RerouteConnections**         | 此方法将重排连接在指定形状上的所有连接符；如果指定的形状是连接符，就重排该连接符。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleHeight**                | 按指定的比例调整形状的高度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整高度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleWidth**                 | 按指定的比例调整形状的宽度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整宽度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                     | 选择对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetShapesDefaultProperties** | 将指定形状的格式设置为形状的默认格式。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup**                    | 取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。返回值：一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象，它代表取消组合的形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**                     | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Adjustments**         | 返回一个 **Adjustments** 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlternativeText**     | 返回或设置一个当 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象保存为网页时，该对象的描述性（可选）文本字符串。**String** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShapeType**       | 返回或设置指定的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)或 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoAutoShapeType%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**     | 返回或设置背景样式。可读/写 **MsoBackgroundStyleIndex** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BlackWhiteMode**      | 返回或设置一个值，该值指明以黑白模式查看演示文稿时指定形状出现的方式。**MsoBlackWhiteMode**，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Callout**             | 返回一个 **CalloutFormat** 对象，它包含指定形状的标注格式属性。应用于代表线形标注的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**               | 返回 **Chart** 对象，该对象代表形状区域中包含的图表。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Child**               | 如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 **msoTrue**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConnectionSiteCount** | 返回指定形状中的连接部位的数量。**Long** 型，只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connector**           | 如果指定的形状是连接符，则此属性为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConnectorFormat**     | 返回一个包含连接符格式属性的 **ConnectorFormat** 对象。应用于代表连接符的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**               | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                | 返回指定形状的 **FillFormat** 对象或指定图表的 **ChartFillFormat** 对象，这些对象包含形状或图表的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                | 返回一个包含形状区域发光格式属性的指定形状区域的 **GlowFormat** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**          | 返回一个 **GroupShapes** 对象，它代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Item** 方法可从形状组中返回单个形状。应用于代表成组形状的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChart**            | 返回形状区域是否包含图表。**MsoTriState** 类型，只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalFlip**      | 如果指定的形状绕水平对称轴翻转，则为 **True**。**MsoTriState**，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                  | 返回一个 Long 值，它代表指定对象的类型。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                | 返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**                | 返回一个 **LineFormat** 对象，它包含指定形状的线条格式属性。对于线条，**LineFormat** 对象代表线条本身；而对于带有边框的形状，**LineFormat** 对象代表边框。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAspectRatio**     | 如果指定的形状在调整大小时其原始比例保持不变，则此属性为 **True**。如果调整大小时可以分别更改形状的高度和宽度，则此属性为 **False**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nodes**               | 返回一个 **ShapeNodes** 集合，它代表指定形状的几何描述。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentGroup**         | 返回 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，它代表子形状或子形状区域的共同父形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**       | 返回一个 **PictureFormat** 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reflection**          | 返回指定形状区域的 **ReflectionFormat** 对象，该对象包含形状区域的映像格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rotation**            | 返回或设形状的旋转角度（以度为单位）。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**              | 返回一个只读的 **ShadowFormat** 对象，它包含指定形状的阴影格式属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeStyle**          | 返回或设置 **MsoShapeStyleIndex**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoShapeStyleIndex%20%E6%9E%9A%E4%B8%BE.html)类型的值，该值代表形状区域的形状样式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoftEdge**            | 返回指定形状区域的 **SoftEdgeFormat** 对象，该对象包含形状区域的柔化边缘格式设置属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**          | 该属性返回一个 **TextEffectFormat** 对象，它包含指定形状的文本效果格式属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame**           | 返回一个 **TextFrame** 对象，它包含指定形状的对齐方式和定位属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**          | 返回一个 **TextFrame2** 对象，该对象包含指定形状区域的文本格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**              | 返回一个 **ThreeDFormat** 对象，它包含指定形状的三维效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**               | 返回或设置与指定的形状区域关联的可选文本的标题。可读/写。返回String值 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                 | 返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                | 返回一个代表形状类型的 **MsoShapeType** 值。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalFlip**        | 如果指定的形状绕垂直坐标轴翻转，则此属性为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Vertices**            | 将指定任意多边形形状的顶点（及贝塞尔曲线的控制点）坐标作为一系列坐标对返回。可将此属性返回的数组用作 **AddCurve** 方法或 **AddPolyLine** 方法的参数。**Variant** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**             | 返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrderPosition**      | 返回指定形状在 z-顺序中的位置。**Long** 型，只读。           |

**成员方法**

#### **ShapeRange.Align**

对齐指定形状区域中的形状。

**语法**

**express.Align(AlignCmd, RelativeTo)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**    | **说明**                               |
| ------------ | ------------- | --------------- | -------------------------------------- |
| *AlignCmd*   | 必选          | **MsoAlignCmd** | 指定某个指定形状范围中形状的对齐方式。 |
| *RelativeTo* | 必选          | **MsoTriState** | 不在 ET 中使用。必须为 False。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中指定范围的所有形状的左边与该范围内最左边的形状的左边对齐。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.SelectAll()     Selection.ShapeRange.Align(msoAlignLefts, false) }` |

#### **ShapeRange.Apply**

应用通过 **PickUp**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.PickUp)方法复制的指定形状格式。

**语法**

**express.Apply()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/   let myDocument = Application.Worksheets.Item(1)   myDocument.Shapes.Item(1).PickUp()   myDocument.Shapes.Item(2).Apply() }` |

#### **ShapeRange.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Distribute**

水平或垂直地分布指定的形状区域中的各形状。

**语法**

**express.Distribute(DistributeCmd, RelativeTo)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**         | **说明**                                   |
| --------------- | ------------- | -------------------- | ------------------------------------------ |
| *DistributeCmd* | 必选          | **MsoDistributeCmd** | 指定该范围内的形状是水平分布还是垂直分布。 |
| *RelativeTo*    | 必选          | **MsoTriState**      | 不在 ET 中使用。必须为 False。             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例在 myDocument 上定义了一个包含所有自选形状对象的形状子集，然后水平地分布该子集中的形状。最左边的形状将保留在原位。*/   let myDocument = Application.Worksheets.Item(1)   let myShapes = myDocument.Shapes   let numAutoShapes   let autoShpArray = []   let numShapes = myShapes.Count      if(numShapes > 1) {       numAutoShapes = 0       for(i = 1 ;i <= numShapes; i++){           if(myShapes.Item(i).Type == msoAutoShape) {               numAutoShapes++               autoShpArray[numAutoShapes] = myShapes.Item(i).Name           }       }       if(numAutoShapes > 1) {           let asRange = myShapes.Range(autoShpArray)           asRange.Distribute(msoDistributeHorizontally, false)       }   } }` |

#### **ShapeRange.Duplicate**

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate()**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Flip**

绕指定形状的水平或垂直对称轴翻转该形状。

**语法**

**express.Flip(FlipCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**   | **说明**                         |
| --------- | ------------- | -------------- | -------------------------------- |
| *FlipCmd* | 必选          | **MsoFlipCmd** | 指定形状是水平翻转还是垂直翻转。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例向 myDocument 中添加三角形，复制该三角形，然后垂直翻转所复制的三角形，并将其设置为红色。*/   let myDocument = Application.Worksheets.Item(1)   let shape = myDocument.Shapes.AddShape(msoShapeRightTriangle, 10, 10, 50, 50).Duplicate()   shape.Fill.ForeColor.RGB = (255, 0, 0)   shape.Flip(msoFlipVertical) }` |

#### **ShapeRange.Group**

将指定区域中的形状形成一组。

返回值：一个代表组合形状的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

**语法**

**express.Group()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于一组形状作为单个形状处理，所以创建和分解形状组将更改 [**Shapes** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

[**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象必须是数据透视表字段的数据区域中的单个单元格。如果试图对多个单元格应用该方法，将会失败（不显示错误消息）。

#### **ShapeRange.IncrementLeft**

以指定磅数水平移动指定形状。

**语法**

**express.IncrementLeft(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 以磅为单位指定形状水平移动的距离。正值使形状向右移动，负值使形状向左移动。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/   let myDocument = Application.Worksheets.Item(1)   let duplicate = myDocument.Shapes.Item(1).Duplicate()   duplicate.Fill.PresetTextured(msoTextureGranite)   duplicate.IncrementLeft(70)   duplicate.IncrementTop(-50)   duplicate.IncrementRotation(30) }` |

#### **ShapeRange.IncrementRotation**

使指定的形状按指定度数值绕 Z 轴旋转。使用 **Rotation**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Rotation)属性可设置形状的绝对转角。

**语法**

**express.IncrementRotation(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状的水平旋转量，以度为单位。正值为顺时针旋转形状，负值逆时针旋转形状。 |

**说明**

如果要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 **IncrementRotationX** 方法或 **IncrementRotationY** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/   let myDocument = Application.Worksheets.Item(1)   let duplicate = myDocument.Shapes.Item(1).Duplicate()   duplicate.Fill.PresetTextured(msoTextureGranite)   duplicate.IncrementLeft(70)   duplicate.IncrementTop(-50)   duplicate.IncrementRotation(30) }` |

#### **ShapeRange.IncrementTop**

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状对象垂直移动的距离，以磅为单位。正值将形状下移，负值将形状上移。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/   let myDocument = Worksheets.Item(1)   let duplicate = myDocument.Shapes.Item(1).Duplicate()   duplicate.Fill.PresetTextured(msoTextureGranite)   duplicate.IncrementLeft(70)   duplicate.IncrementTop(-50)   duplicate.IncrementRotation(30) }` |

#### **ShapeRange.Item**

从集合中返回一个对象。

返回值：包含在集合中的一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     /*本示例对某个形状区域中第二个形状的 OnAction 属性进行设置。如果变量 sr 不代表 ShapeRange 对象，则本示例无效。*/     let sr = Application.Worksheets.Item(1).Shapes.Range([1, 3])     sr.Item(2).OnAction = "ShapeAction" }` |

#### **ShapeRange.PickUp**

复制指定形状的格式。使用 **Apply**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Apply)方法可将复制的格式应用到其他形状。

**语法**

**express.PickUp()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/   let myDocument = Application.Worksheets.Item(1)   myDocument.Shapes.Item(1).PickUp()   myDocument.Shapes.Item(2).Apply() }` |

#### **ShapeRange.Regroup**

将指定形状原先所属的形状区域重新组合。以单一的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象返回重新组合的形状。

返回值类型：Shape

**语法**

**express.Regroup()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

**Regroup** 方法在指定的 **ShapeRange** 集合中查找原先曾被组合的形状，并仅还原所找到的第一个形状原来所属的形状组。因此，如果指定的形状子集所包含的各个图形原来曾属于不同的图形组，就只会恢复其中的一个图形组。

请注意，由于一组形状作为单个形状处理，因此将形状组合及取消组合都会改变 [**Shapes** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合中的项目数，而且还会改变集合中受影响的项目后面的项目索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对当前活动窗口中选定的形状进行重新分组。如果这些形状并未进行过组合和取消组合操作，则本示例将失效。*/ Application.ActiveWindow.Selection.ShapeRange.Regroup()` |

#### **ShapeRange.RerouteConnections**

此方法将重排连接在指定形状上的所有连接符；如果指定的形状是连接符，就重排该连接符。

**语法**

**express.RerouteConnections()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

重排连接符使其以最短的路径连接形状。重排时，**RerouteConnections** 方法可能会断开连接符的两端并将其重新连接到形状的其他位置。

如果该方法应用于一个连接符，则只重排该连接符；如果该方法应用于一个已连接的形状，则重排该形状上所有的连接符。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例将两个矩形添加到 myDocument，用曲线连接符连接两个矩形，然后重排连接符使两个矩形间采用最短的路径。请注意，RerouteConnections 方法用于调整连接符的大小和位置并决定要连接到哪个连接结点，因此，最初为 BeginConnect 和 EndConnect 方法的 ConnectionSite 参数指定的值是不相关的。*/   let myDocument = ApplicationWorksheets.Item(1)   let s = myDocument.Shapes   let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)   let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)   let newConnector = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)   newConnector.ConnectorFormat.BeginConnect(firstRect, 1)   newConnector.ConnectorFormat.EndConnect(secondRect, 1)   newConnector.RerouteConnections() }` |

#### **ShapeRange.ScaleHeight**

按指定的比例调整形状的高度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整高度。

**语法**

**express.ScaleHeight(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------------------ | ------------- | --------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**      | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数设为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState** | 如果为 msoTrue，则相对于形状的原有尺寸来调整高度。如果该值为 msoFalse，则相对于形状的当前尺寸来调整高度。仅当指定的形状是图片或 OLE 对象时，才能将此参数指定为 msoTrue。 |
| *Scale*                  | 可选          | **Variant**     | MsoScaleFrom 的常量之一，它指定调整形状大小时，该形状哪一部分的位置将保持不变。 |

**说明**

| **MsoTriState** 可以是下列 [**MsoTriState** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoTriState%20%E6%9E%9A%E4%B8%BE.html)常量之一。 |
| ------------------------------------------------------------ |
| **msoCTrue**。不应用于此属性。                               |
| **msoFalse**。相对于形状的当前尺寸来调整形状的大小。         |
| **msoTriStateMixed**。不应用于此属性。                       |
| **msoTriStateToggle**。不应用于此属性。                      |
| **msoTrue**。相对于形状的初始尺寸来调整形状的大小。          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/   let myDocument = Application.Worksheets.Item(1)   let s = myDocument.Shapes		   for(let i = 1; i <= s.Count; i++) {       switch(s.Item(i).Type) {       case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:           s.Item(i).ScaleHeight(1.75, msoTrue)           s.Item(i).ScaleWidth(1.75, msoTrue)           break       default:           s.Item(i).ScaleHeight(1.75, msoFalse)           s.Item(i).ScaleWidth(1.75, msoFalse)           break   	}   } }` |

#### **ShapeRange.ScaleWidth**

按指定的比例调整形状的宽度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整宽度。

**语法**

**express.ScaleWidth(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------------------ | ------------- | --------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**      | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数设为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState** | 如果为 False，则相对于形状的原有尺寸来调整宽度。仅当指定的形状是图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选          | **Variant**     | MsoScaleFrom 的常量之一，它指定调整形状大小时，该形状哪一部分的位置将保持不变。 |

**说明**

| **MsoTriState** 可为以下 [**MsoTriState** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoTriState%20%E6%9E%9A%E4%B8%BE.html)常量之一。 |
| ------------------------------------------------------------ |
| **msoCTrue**。不应用于此属性。                               |
| **msoFalse**。相对于形状的当前尺寸来调整形状的大小。         |
| **msoTriStateMixed**。不应用于此属性。                       |
| **msoTriStateToggle**。不应用于此属性。                      |
| **msoTrue**。仅当指定的形状是图片或 OLE 对象时，才能使用此参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/   let myDocument = Application.Worksheets.Item(1)   let s = myDocument.Shapes		   for(let i = 1; i <= s.Count; i++) {       switch(s.Item(i).Type) {       case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:           s.Item(i).ScaleHeight(1.75, msoTrue)           s.Item(i).ScaleWidth(1.75, msoTrue)           break       default:           s.Item(i).ScaleHeight(1.75, msoFalse)           s.Item(i).ScaleWidth(1.75, msoFalse)           break   	}   } }` |

#### **ShapeRange.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

#### **ShapeRange.SetShapesDefaultProperties**

将指定形状的格式设置为形状的默认格式。

**语法**

**express.SetShapesDefaultProperties()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例向 myDocument 中添加矩形，设置该矩形的填充样式，将矩形的格式设置为该矩形的默认格式，然后在文档中添加另一个较小的矩形。第二个矩形的填充格式与第一个矩形一样。*/   let myDocument = Application.Worksheets.Item(1)   let s = myDocument.Shapes   let a =s.AddShape(msoShapeRectangle, 5, 5, 80, 60)   let f = a.Fill   f.ForeColor.RGB = (0, 0, 255)   f.BackColor.RGB = (0, 204, 255)   f.Patterned(msoPatternHorizontalBrick)      // Set formatting as default formatting   a.SetShapesDefaultProperties()      // Create new shape with default formatting   s.AddShape(msoShapeRectangle, 90, 90, 40, 30) }` |

#### **ShapeRange.Ungroup**

取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。

返回值：一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象，它代表取消组合的形状。

**语法**

**express.Ungroup()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于一组形状被视为单个对象，因此对形状进行组合或取消组合操作时，将会更改 **Shapes**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例取消 myDocument 中的所有形状组合，并分解所有的图片和 OLE 对象。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     let s = myDocument.Shapes     for(let i = 1; i <= s.Count; i++) {         s.Item(i).Ungroup()     } }  /*此示例取消 myDocument 中的所有形状组合，但并不取消文档中图片和 OLE 对象的组合。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     let s = myDocument.Shapes     for(let i = 1; i <= s.Count; i++) {         if(s.Item(i).Type == msoGroup) {             s.Item(i).Ungroup()         }     } }` |

#### **ShapeRange.ZOrder**

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

**语法**

**express.ZOrder(ZOrderCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                   |
| ----------- | ------------- | ---------------- | ------------------------------------------ |
| *ZOrderCmd* | 必选          | **MsoZOrderCmd** | 指定根据其他形状将指定的形状移到什么位置。 |

**说明**

| **MsoZOrderCmd** 可为以下 [**MsoZOrderCmd** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoZOrderCmd%20%E6%9E%9A%E4%B8%BE.html)常量之一。 |
| ------------------------------------------------------------ |
| **msoBringForward**                                          |
| **msoBringInFrontOfText**。仅用于 WPS。                      |
| **msoBringToFront**                                          |
| **msoSendBackward**                                          |
| **msoSendBehindText**。仅用于 WPS。                          |
| **msoSendToBack**                                            |

使用 **ZOrderPosition** 属性可确定形状在 z-次序中的当前位置。

**示例**

**成员属性**

#### **ShapeRange.Adjustments**

返回一个 **Adjustments** 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。

**语法**

**express.Adjustments**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.AlternativeText**

返回或设置一个当 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象保存为网页时，该对象的描述性（可选）文本字符串。**String** 型，可读写。

**语法**

**express.AlternativeText**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

可选文字可以显示在 Web 浏览器中某个形状的图像位置上，或者当鼠标指针停留在图像上时，这些文本将直接显示在图像之上（在支持这些功能的浏览器中）。

#### **ShapeRange.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     /*本示例显示一条有关创建 myObject 的应用程序的消息。*/     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **ShapeRange.AutoShapeType**

返回或设置指定的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)或 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoAutoShapeType%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。

**语法**

**express.AutoShapeType**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **ConnectorFormat** 对象的 **Type** 属性设置或返回连接符类型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中的所有 16 角星都替换为 32 角星。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     for(let i=1;i <= myDocument.Shapes.Count;i++) {         if(myDocument.Shapes.Item(i).AutoShapeType == msoShape16pointStar) {             myDocument.Shapes.Item(i).AutoShapeType = msoShape32pointStar         }     } }` |

#### **ShapeRange.BackgroundStyle**

返回或设置背景样式。可读/写 **MsoBackgroundStyleIndex** 类型。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.BlackWhiteMode**

返回或设置一个值，该值指明以黑白模式查看演示文稿时指定形状出现的方式。**MsoBlackWhiteMode**，可读写。

**语法**

**express.BlackWhiteMode**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例使 wksOne 上的第一个形状以黑白模式显示。当用户以黑白模式查看演示文稿时，第一个形状为黑色，而不管它在颜色模式中为何种颜色。*/ function UseBlackWhiteMode() {     let wksOne = Application.Worksheets.Item(1)     wksOne.Shapes.Item(1).BlackWhiteMode = msoBlackWhiteGrayOutline }` |

#### **ShapeRange.Callout**

返回一个 **CalloutFormat** 对象，它包含指定形状的标注格式属性。应用于代表线形标注的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。

**语法**

**express.Callout**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加椭圆和指向该椭圆的标注。该标注的文字没有边框，但用垂直的强调线分开标注文字和标注线。*/ function test(){   let myDocument = Application.Worksheets.Item(1)   let myShape = myDocument.Shapes   myShape.AddShape(msoShapeOval, 180, 200, 280, 130)   let Add = myShape.AddCallout(msoCalloutTwo, 420, 170, 170, 40)   Add.TextFrame.Characters().Text = "My oval"   let Call = Add.Callout   Call.Accent = true   Call.Border = false }` |

#### **ShapeRange.Chart**

返回 **Chart** 对象，该对象代表形状区域中包含的图表。只读。

**语法**

**express.Chart**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Child**

如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 **msoTrue**。**MsoTriState** 类型，只读。

**语法**

**express.Child**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

| **MsoTriState** 可为以下 **MsoTriState** 常量之一:          |
| ----------------------------------------------------------- |
| **msoCTrue**。不应用于此属性。                              |
| 如果选择的形状不是子形状，则为 **msoFalse**。               |
| 如果只有一些选择的形状是子形状，则为 **msoTriStateMixed**。 |
| **msoTriStateToggle**。不应用于此属性。                     |
| 如果选择的形状是子形状，则为 **msoTrue**。                  |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择画布中的第一个形状，并且如果选择的形状是子形状，则用指定的颜色填充该形状。本示例假定活动工作表上的一个绘图画布中包含多个形状。*/ function FillChildShape() {     //Select the first shape in the drawing canvas.     Application.ActiveSheet.Shapes.Item(1).CanvasItems(1).Select()     //Fill selected shape if it is a child shape.     if(Selection.ShapeRange.Child == msoTrue) {         Selection.ShapeRange.Fill.ForeColor.RGB = (100, 0, 200)     }     else {         MsgBox("This shape is not a child shape.")     } }` |

#### **ShapeRange.ConnectionSiteCount**

返回指定形状中的连接部位的数量。**Long** 型，只读。

**语法**

**express.ConnectionSiteCount**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例向 myDocument 中添加两个矩形，并且用两个连接符连接这两个矩形。两个连接符的起始点都连到第一个矩形上的第一个连接部位；连接符的终止点分别连到第二个矩形上的第一个和最后一个连接部位上。*/   let myDocument = Application.Worksheets.Item(1)   let s = myDocument.Shapes   let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)   let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)   let lastsite = secondRect.ConnectionSiteCount   let Con1 = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat   Con1.BeginConnect(firstRect, 1)   Con1.EndConnect(secondRect, 1)   let Con2 = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat   Con2.BeginConnect(firstRect, 1)   Con2.EndConnect(secondRect, lastsite) }` |

#### **ShapeRange.Connector**

如果指定的形状是连接符，则此属性为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.Connector**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例删除了 myDocument 中所有的连接符。*/   let myDocument = Application.Worksheets.Item(1)   let myShape = myDocument.Shapes   for(let i=myShape.Count;i >= 1;i--) {        if(myShape.Item(i).Connector) {           myShape.Item(i).Delete()       }   } }` |

#### **ShapeRange.ConnectorFormat**

返回一个包含连接符格式属性的 **ConnectorFormat** 对象。应用于代表连接符的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。

**语法**

**express.ConnectorFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例向 myDocument 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/   let myDocument = Application.Worksheets.Item(1)   let s = myDocument.Shapes   let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)   let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)   let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)   let Con = c.ConnectorFormat   Con.BeginConnect(firstRect, 1)   Con.EndConnect(secondRect, 1)   c.RerouteConnections()   Con.BeginDisconnect()   Con.EndDisconnect() }` |

#### **ShapeRange.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ShapeRange.Fill**

返回指定形状的 **FillFormat** 对象或指定图表的 **ChartFillFormat** 对象，这些对象包含形状或图表的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加矩形，然后设置该矩形的填充格式的前景色、背景色和渐变。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     let myShape = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill     myShape.ForeColor.RGB = (128, 0, 0)     myShape.BackColor.RGB = (170, 170, 170)     myShape.TwoColorGradient(msoGradientHorizontal, 1) }` |

#### **ShapeRange.Glow**

返回一个包含形状区域发光格式属性的指定形状区域的 **GlowFormat** 对象。只读。

**语法**

**express.Glow**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

发光效果为图形添加了色彩鲜明的彩色边缘。

#### **ShapeRange.GroupItems**

返回一个 **GroupShapes** 对象，它代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Item** 方法可从形状组中返回单个形状。应用于代表成组形状的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例向 myDocument 中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色。*/   let myDocument = Application.Worksheets.Item(1)   let myShape = myDocument.Shapes   myShape.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"   myShape.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"   myShape.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"   let myGroup = myShape.Range(["shpOne", "shpTwo", "shpThree"]).Group()   myGroup.Fill.PresetTextured(msoTextureBlueTissuePaper)   myGroup.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble) }` |

#### **ShapeRange.HasChart**

返回形状区域是否包含图表。**MsoTriState** 类型，只读。

**语法**

**express.HasChart**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Height**

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.HorizontalFlip**

如果指定的形状绕水平对称轴翻转，则为 **True**。**MsoTriState**，只读。

**语法**

**express.HorizontalFlip**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `funcion test(){   /*如果 myDocument 中的形状进行过水平或垂直翻转，则此示例将恢复其原始状态。*/   let myDocument = Application.Worksheets.Item(1)   for(let i=1;i <= myDocument.Shapes.Count;i++) {       if(myDocument.Shapes.Item(i).HorizontalFlip) {           myDocument.Shapes.Item(i).Flip(msoFlipHorizontal)       }       if(myDocument.Shapes.Item(i).VerticalFlip) {           myDocument.Shapes.Item(i).Flip(msoFlipVertical)       }   } }` |

#### **ShapeRange.ID**

返回一个 Long 值，它代表指定对象的类型。

**语法**

**express.ID**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

#### **ShapeRange.Left**

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Line**

返回一个 **LineFormat** 对象，它包含指定形状的线条格式属性。对于线条，**LineFormat** 对象代表线条本身；而对于带有边框的形状，**LineFormat** 对象代表边框。只读。

**语法**

**express.Line**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例将一条蓝色虚线添至 myDocument。*/   let myDocument = Application.Worksheets.Item(1)   let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line   myLine.DashStyle = msoLineDashDotDot   myLine.ForeColor.RGB = (50, 0, 128)      /*此示例向 myDocument 添加一个十字形，然后将其边框设为红色、8 磅宽。*/   let myDocument = Application.Worksheets.Item(1)   let myLine = myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line   myLine.Weight = 8   myLine.ForeColor.RGB = 255, 0, 0    }` |

#### **ShapeRange.LockAspectRatio**

如果指定的形状在调整大小时其原始比例保持不变，则此属性为 **True**。如果调整大小时可以分别更改形状的高度和宽度，则此属性为 **False**。**MsoTriState** 类型，可读写。

**语法**

**express.LockAspectRatio**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

| **MsoTriState** 可为以下 **MsoTriState** 常量之一:      |
| ------------------------------------------------------- |
| **msoCTrue**                                            |
| **msoFalse**。调整大小时可以分别更改形状的高度和宽度。  |
| **msoTriStateMixed**                                    |
| **msoTriStateToggle**                                   |
| **msoTrue**。指定的形状在调整大小时其原始比例保持不变。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例向 myDocument 中添加一个立方体。可移动并调整该立方体，但不能重新设置其比例。*/   let myDocument = Application.Worksheets.Item(1)   myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200).LockAspectRatio = msoTrue }` |

#### **ShapeRange.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Nodes**

返回一个 **ShapeNodes** 集合，它代表指定形状的几何描述。

**语法**

**express.Nodes**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

此属性适用于代表任意多边形的 **Shape** 或 **ShapeRange** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例在 myDocument 中第三个形状的第四个结点后添加一个带有一段曲线的平滑结点。第三个形状必须是至少有四个结点的任意多边形。*/   let myDocument = Application.Worksheets.Item(1)   let myNode = myDocument.Shapes.Item(3).Nodes   myNode.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100) }` |

#### **ShapeRange.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ParentGroup**

返回 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，它代表子形状或子形状区域的共同父形状。

**语法**

**express.ParentGroup**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在此示例中，ET 向活动工作表添加两个形状，然后通过删除组合的父形状来删除这两个形状。*/ function ParentGroup() {     let myShape = Application.ActiveSheet.Shapes     myShape.AddShape(1, 10, 10, 100, 100)     myShape.AddShape(2, 110, 120, 100, 100)     myShape.Range([1, 2]).Group()     // Using the child shape in the group get the Parent shape.     let pgShape = ActiveSheet.Shapes.Item(1).GroupItems.Item(1).ParentGroup     alert("The two shapes will now be deleted.")     // Delete the parent shape.     pgShape.Delete() }` |

#### **ShapeRange.PictureFormat**

返回一个 **PictureFormat** 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。只读。

**语法**

**express.PictureFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/   let myDocument = Application.Worksheets.Item(1)   let Pic = myDocument.Shapes.Item(1).PictureFormat   Pic.Brightness = 0.3   Pic.Contrast = .75 }` |

#### **ShapeRange.Reflection**

返回指定形状区域的 **ReflectionFormat** 对象，该对象包含形状区域的映像格式属性。只读。

**语法**

**express.Reflection**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Rotation**

返回或设形状的旋转角度（以度为单位）。**Single** 型，可读写。

**语法**

**express.Rotation**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

#### **ShapeRange.Shadow**

返回一个只读的 **ShadowFormat** 对象，它包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ShapeStyle**

返回或设置 **MsoShapeStyleIndex**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoShapeStyleIndex%20%E6%9E%9A%E4%B8%BE.html)类型的值，该值代表形状区域的形状样式。可读/写。

**语法**

**express.ShapeStyle**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.SoftEdge**

返回指定形状区域的 **SoftEdgeFormat** 对象，该对象包含形状区域的柔化边缘格式设置属性。只读。

**语法**

**express.SoftEdge**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.TextEffect**

该属性返回一个 **TextEffectFormat** 对象，它包含指定形状的文本效果格式属性。

**语法**

**express.TextEffect**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例中，如果 myDocument 的第三个形状为艺术字，则将它的字体样式设置为加粗。*/   let myDocument = Application.Worksheets.Item(1)   let myShape = myDocument.Shapes.Item(3)   if(myShape.Type == msoTextEffect) {       myShape.TextEffect.FontBold = true   } }` |

#### **ShapeRange.TextFrame**

返回一个 **TextFrame** 对象，它包含指定形状的对齐方式和定位属性。

**语法**

**express.TextFrame**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将形状一中文本框内的文本设置为两端对齐。如果形状一中不包含文本框，则此示例无效。*/ Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.HorizontalAlignment = xlHAlignJustify` |

#### **ShapeRange.TextFrame2**

返回一个 **TextFrame2** 对象，该对象包含指定形状区域的文本格式。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ThreeD**

返回一个 **ThreeDFormat** 对象，它包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本例设置应用于 myDocument 中第一个形状的三维效果：深度、延伸色、延伸方向以及照明方向。*/   let myDocument = Application.Worksheets.Item(1)   let myThree = myDocument.Shapes.Item(1).ThreeD   myThree.Visible = true   myThree.Depth = 50   myThree.ExtrusionColor.RGB = (255, 100, 255)   // RGB value for purple   myThree.SetExtrusionDirection(msoExtrusionTop)   myThree.PresetLightingDirection = msoLightingLeft }` |

#### **ShapeRange.Title**

返回或设置与指定的形状区域关联的可选文本的标题。可读/写。

返回String值

**语法**

**express.Title**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

若要设置形状区域的可选文本的描述性文本字符串，请使用 **AlternativeText** 属性。

#### **ShapeRange.Top**

返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Type**

返回一个代表形状类型的 **MsoShapeType** 值。

**语法**

**express.Type**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.VerticalFlip**

如果指定的形状绕垂直坐标轴翻转，则此属性为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*如果 myDocument 中的形状进行过水平或垂直翻转，则本示例将恢复其原始状态。*/   let myDocument = Application.Worksheets.Item(1)   for(let i=1;i <= myDocument.Shapes.Count;i++) {       if(myDocument.Shapes.Item(i).HorizontalFlip) {           myDocument.Shapes.Item(i).Flip(msoFlipHorizontal)       }       if(myDocument.Shapes.Item(i).VerticalFlip) {           myDocument.Shapes.Item(i).Flip(msoFlipVertical)       }   } }` |

#### **ShapeRange.Vertices**

将指定任意多边形形状的顶点（及贝塞尔曲线的控制点）坐标作为一系列坐标对返回。可将此属性返回的数组用作 **AddCurve** 方法或 **AddPolyLine** 方法的参数。**Variant** 型，只读。

**语法**

**express.Vertices**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

下表显示 **Vertices** 属性数组 `vertArray()` 中的值如何与三角形的 vertices 坐标相关联。

| vertArray 元素    | 内容                                   |
| ----------------- | -------------------------------------- |
| `vertArray(1, 1)` | 第一个顶点与文档的左边界之间的水平距离 |
| `vertArray(1, 2)` | 第一个顶点与文档的顶端之间的垂直距离   |
| `vertArray(2, 1)` | 第二个顶点与文档的左边界之间的水平距离 |
| `vertArray(2, 2)` | 第二个顶点与文档的顶端之间的垂直距离   |
| `vertArray(3, 1)` | 第三个顶点与文档的左边界之间的水平距离 |
| `vertArray(3, 2)` | 第三个顶点与文档的顶端之间的垂直距离   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例将 myDocument 上第一个形状的顶点坐标分配给数组变量 vertArray()，并显示第一个顶点的坐标。*/   let myDocument = Application.Worksheets.Item(1)   let myShape = myDocument.Shapes.Item(1)   let vertArray = [myShape.Vertices]   x1 = vertArray[1, 1]   y1 = vertArray[1, 2]   alert("First vertex coordinates: " + x1 + ", " + y1)    } function test(){   /*此示例创建一个曲线，该曲线的几何说明与myDocument中第一个形状相同。此示例要求第一个形状必须包含3n+1个顶点。*/   let myDocument = Worksheets.Item(1)   let myShape = myDocument.Shapes   myShape.AddCurve(myShape.Item(1).Vertices)    }` |

#### **ShapeRange.Visible**

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Width**

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ZOrderPosition**

返回指定形状在 z-顺序中的位置。**Long** 型，只读。

**语法**

**express.ZOrderPosition**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

若要设置形状在 z-顺序中的位置，请用 **ZOrder** 方法。

形状在 z-顺序中的位置对应于在 [**Shapes** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合中的索引号。例如，如果在 `myDocument` 中有四个形状，则 `myDocument.Shapes(1)` 表达式返回 z-顺序中最后一个形状，表达式 `myDocument.Shapes(4)` 返回 z-顺序中第一个形状。

每次向集合中添加一个新的形状时，默认情况下它会被添加到 z-次序的最前端。

**示例**

适用环境：web

适用平台：windows/linux