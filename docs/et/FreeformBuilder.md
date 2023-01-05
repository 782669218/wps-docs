**FreeformBuilder**



代表任意多边形创建时的几何属性。

**说明**

使用 **BuildFreeform** 方法可返回 **FreeformBuilder** 对象。使用 **AddNodes** 方法可在任意多边形中添加节点。使用 **ConvertToShape** 方法可创建 **FreeformBuilder** 对象中定义的形状，并将它添加到 **Shapes** 集合中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例在 myDocument 中添加一个具有四段的任意多边形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)     shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)     shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 2000)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)     shape1.ConvertToShape() }` |

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddNodes**       | 在当前形状中添加一点，然后绘制一个从当前节点到添加的最后一个节点的线条。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToShape** | 创建一个具有指定 **FreeformBuilder** 对象的几何特性的形状。返回一个代表新形状的 **Shape** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **FreeformBuilder.AddNodes**

在当前形状中添加一点，然后绘制一个从当前节点到添加的最后一个节点的线条。

**语法**

**express.AddNodes(SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3)**

*express*   一个代表 **FreeformBuilder** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**       | **说明**                                                     |
| ------------- | ------------- | ------------------ | ------------------------------------------------------------ |
| *SegmentType* | 必选          | **MsoSegmentType** | 要添加的线段的类型。                                         |
| *EditingType* | 必选          | **MsoEditingType** | 顶点的编辑属性。                                             |
| *X1*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。 如果新节点的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第一个控制点的水平距离（以磅为单位）。 |
| *Y1*          | 必选          | **Single**         | 如果新线段的 EditingType 为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。 如果新节点的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第一个控制点的水平距离（以磅为单位）。 |
| *X2*          | 可选          | **Variant**        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。 |
| *Y2*          | 可选          | **Variant**        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。 |
| *X3*          | 可选          | **Variant**        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。 |
| *Y3*          | 可选          | **Variant**        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。 |

**说明**

| **MsoSegmentType** 可以是下列 **MsoSegmentType** 常量之一。 |
| ----------------------------------------------------------- |
| **msoSegmentLine**                                          |
| **msoSegmentCurve**                                         |

| **MsoEditingType** 可以是下列 **MsoEditingType** 常量之一。  |
| ------------------------------------------------------------ |
| **msoEditingAuto**                                           |
| **msoEditingCorner**                                         |
| 不能是 **msoEditingSmooth** 或 **msoEditingSymmetric**如果 *SegmentType* 为 **msoSegmentLine**，那么 *EditingType* 就必须是 **msoEditingAuto**。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加带有四条线段的任意多边形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)     shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)     shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)     shape1.ConvertToShape() }` |

#### **FreeformBuilder.ConvertToShape**

创建一个具有指定 **FreeformBuilder** 对象的几何特性的形状。返回一个代表新形状的 **Shape** 对象。

**语法**

**express.ConvertToShape()**

*express*   一个代表 **FreeformBuilder** 对象的变量。

**返回值**

Shape

**说明**

对于 **FreeformBuilder** 对象，至少要应用一次 **AddNodes** 方法，然后才能使用 **ConvertToShape** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个有五个顶点的任意多边形添加到 myDocument 中。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)     shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)     shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)     shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)     shape1.ConvertToShape() }` |

**成员属性**

#### **FreeformBuilder.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **FreeformBuilder** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **FreeformBuilder.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FreeformBuilder** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **FreeformBuilder.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FreeformBuilder** 对象的变量。

适用环境：web

适用平台：windows/linux