**Shape**



代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

**说明**

**Shape** 对象是 **Shapes** 集合的成员。**Shapes** 集合包含某个工作簿中的所有形状。

**注释：**有三个代表形状的对象：**Shapes** 集合，它代表工作簿中所有的形状；[**ShapeRange**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange) 集合，它代表工作簿中形状的指定子集（例如，**ShapeRange** 对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）；**Shape** 对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用 **ShapeRange** 集合。

使用 Shape 对象

以下各节说明了如何：

- 返回与连接符的端点相连的形状。
- 返回新建的任意多边形。
- 返回组中的单个形状。
- 返回新组成的形状组。
- 返回现有的形状。
- 返回选定区域中的形状。返回与连接符的端点相连的形状

要返回一个代表连接符所连接形状之一的 **Shape** 对象，请使用 **BeginConnectedShape** 或 **EndConnectedShape** 属性。

\1. 返回新建的任意多边形

​    使用 **BuildFreeform** 和 **AddNodes** 方法可定义一个新任意多边形的几何特性，使用 **ConvertToShape** 方法可创建任意多边形并返回代表它的 **Shape** 对象。

**2.** 返回组中的单个形状

​    使用 **GroupItems**(*index* )（其中 *index* 是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的 **Shape** 对象。

**3.** 返回新组成的形状组

​    使用 **Group** 或 **Regroup** 方法可将一系列形状分成一组并返回一个 **Shape** 对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组

\4. 返回现有的形状

​    使用 **Shapes**(*index* )（其中 *index* 是形状名称或索引号）可返回代表某个形状的 **Shape** 对象。

\5. 返回选定区域中的形状

​    使用 `Selection.ShapeRange`(*index*)（其中 *index* 是形状名称或索引号）可返回一个代表选定区域中的形状的 **Shape** 对象。

**方法**

|                                                              | 名称        | 说明                                                         |
| ------------------------------------------------------------ | ----------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**   | 应用通过 **PickUp** 方法复制的指定形状格式。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**    | 将对象复制到剪贴板。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**     | 将对象剪切到剪贴板。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**  | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**  | 选择对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup** | 取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。返回值：一个 **ShapeRange** 对象，它代表取消组合的形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**  | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShapeType**   | 返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomRightCell** | 返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**          | 返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**              | 返回一个 Long 值，它代表指定对象的类型。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**            | 返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rotation**        | 返回或设形状的旋转角度（以度为单位）。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**             | 返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopLeftCell**     | 返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回或设置一个代表形状类型的 **MsoShapeType** 值。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**         | 返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**           | 返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。 |

**成员方法**

#### **Shape.Apply**

应用通过 **PickUp** 方法复制的指定形状格式。

**语法**

**express.Apply()**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/ function test() {     Application.Worksheets.Item(1).Shapes.Item(1).PickUp()     Application.Worksheets.Item(1).Shapes.Item(2).Apply() }` |

#### **Shape.Copy**

将对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Cut**

将对象剪切到剪贴板。

**语法**

**express.Cut()**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

#### **Shape.Ungroup**

取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。

返回值：一个 **ShapeRange** 对象，它代表取消组合的形状。

**语法**

**express.Ungroup()**

*express*   一个代表 **Shape** 对象的变量。

**说明**

由于一组形状被视为单个对象，因此对形状进行组合或取消组合操作时，将会更改 **Shapes** 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例取消 myDocument 中的所有形状组合，并分解所有的图片和 OLE 对象。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     for(let i=1;i <= myDocument.Shapes.Count;i++) {         myDocument.Shapes.Item(i).Ungroup()     } }  /*此示例取消 myDocument 中的所有形状组合，但并不取消文档中图片和 OLE 对象的组合。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     for(let i=1;i <= myDocument.Shapes.Count;i++) {         if(myDocument.Shapes.Item(i).Type == msoGroup){             myDocument.Shapes.Item(i).Ungroup()         }     } }` |

#### **Shape.ZOrder**

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

**语法**

**express.ZOrder(ZOrderCmd)**

*express*   一个代表 **Shape** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                   |
| ----------- | ------------- | ---------------- | ------------------------------------------ |
| *ZOrderCmd* | 必选          | **MsoZOrderCmd** | 指定根据其他形状将指定的形状移到什么位置。 |

**说明**

**Shape** 对象是 **Shapes** 集合的成员。**Shapes** 集合包含某个工作簿中的所有形状。

**注释：**有三个代表形状的对象：**Shapes** 集合，它代表工作簿中所有的形状；**ShapeRange** 集合，它代表工作簿中形状的指定子集（例如，**ShapeRange** 对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）；**Shape** 对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用 **ShapeRange** 集合。

使用 Shape 对象

以下各节说明了如何：

- 返回与连接符的端点相连的形状。
- 返回新建的任意多边形。
- 返回组中的单个形状。
- 返回新组成的形状组。
- 返回现有的形状。
- 返回选定区域中的形状。

**1. 返回与连接符的端点相连的形状**

 

​    要返回一个代表连接符所连接形状之一的 **Shape** 对象，请使用 **BeginConnectedShape** 或 **EndConnectedShape** 属性。

\2. 返回新建的任意多边形

​    使用 **BuildFreeform** 和 **AddNodes** 方法可定义一个新任意多边形的几何特性，使用 **ConvertToShape** 方法可创建任意多边形并返回代表它的 **Shape** 对象。

\3. 返回组中的单个形状

​    使用 **GroupItems**(*index* )（其中 *index* 是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的 **Shape** 对象。

\4. 返回新组成的形状组

​    使用 **Group** 或 **Regroup** 方法可将一系列形状分成一组并返回一个 **Shape** 对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组。

\5. 返回现有的形状

使用 **Shapes**(*index* )（其中 *index* 是形状名称或索引号）可返回代表某个形状的 **Shape** 对象。

**6**. 返回选定区域中的形状

使用 `Selection.ShapeRange`(*index*)（其中 *index* 是形状名称或索引号）可返回一个代表选定区域中的形状的 **Shape** 对象。

**示例**

**成员属性**

#### **Shape.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Shape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if(Application.ActiveWorkbook.Application.Value == "ET") {         alert("This is an ET Application object.")     }else {         alert("This is not an ET Application object.")     } }` |

#### **Shape.AutoShapeType**

返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。**MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

*express*   一个代表 **Shape** 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **ConnectorFormat** 对象的 **Type** 属性设置或返回连接符类型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中的所有 16 角星都替换为 32 角星。*/ function test() {     for(let i = 1; i <= Application.Worksheets.Item(1).Shapes.Count; i++) {         if(Application.Worksheets.Item(1).Shapes.Item(i).AutoShapeType = msoShape16pointStar) {             Application.Worksheets.Item(1).Shapes.Item(i).AutoShapeType = msoShape32pointStar         }     } }` |

#### **Shape.BottomRightCell**

返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。

**语法**

**express.BottomRightCell**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Height**

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.ID**

返回一个 Long 值，它代表指定对象的类型。

**语法**

**express.ID**

*express*   一个代表 **Shape** 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

#### **Shape.Left**

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Rotation**

返回或设形状的旋转角度（以度为单位）。**Single** 型，可读写。

**语法**

**express.Rotation**

*express*   一个代表 **Shape** 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

#### **Shape.Top**

返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.TopLeftCell**

返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。

**语法**

**express.TopLeftCell**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Type**

返回或设置一个代表形状类型的 **MsoShapeType** 值。

**语法**

**express.Type**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Visible**

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **Shape** 对象的变量。

#### **Shape.Width**

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **Shape** 对象的变量。

适用环境：web

适用平台：windows/linux