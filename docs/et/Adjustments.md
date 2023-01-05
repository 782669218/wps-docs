**Adjustments**



它包含指定的自选图形、艺术字对象或连接符的调整值的集合。

**说明**

每个调整值代表一种调整控点的调整方法。由于某些调整控点可以按两种方法调整（例如，某些控点既可以水平调整也可以垂直调整），所以形状的调整值数量可以大于调整控点数量。一个形状最多可以有八个调整值。

使用 **Adjustments** 属性可返回 **Adjustments** 对象。使用 **Adjustments**(*index*)（其中 *index* 是调整值的索引号）可返回单个调整值。

不同的形状具有不同数目的调整值，不同类型的调整值在不同的方向上调整形状的几何性质，不同类型的调整值有不同的取值范围。例如，下面的图示显示了右箭头标注的四个调整值各对该标注的几何形状起什么作用。

![具有不同调整控点的右箭头标注](https://qn.cache.wpscdn.cn/encs/doc/office_v13/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/gif/adjlabel_ZA06051188.gif)

| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/close.gif)注释 |
| ------------------------------------------------------------ |
| 由于每个形状有不同的调整值集，校验指定形状的调整行为的最好方法是手动创建一个图例，在打开宏记录器的情况下作调整，然后检查记录的代码。 |

下表概括了不同类型的调整值的有效取值范围。多数情况下，如果指定的调整值超出了有效取值范围，就将用最接近的有效值来代替。

| 调整类型           | 有效值                                                       |
| ------------------ | ------------------------------------------------------------ |
| 线性（水平或垂直） | 通常 0.0 值代表形状的左边界或上边界，而 1.0 值代表形状的右边界或下边界。有效值对应于有效的手动调整。例如，如果只能将调整控点手动拖动形状的一半宽度，则相应的调整值最大为 0.5。对于象连接符和标注这样的形状，0.0 和 1.0 值代表由它们的起始和终止点定义的矩形界限，此时负值和大于 1.0 的值是有效的。 |
| 射线图             | 调整值 1.0 对应于形状宽度。最大值为 0.5，或形状宽度的一半。  |
| 角                 | 值以度表示。如果指定的值超过了 -180 到 180 的范围，则将其折算为该范围内的值。 |

示例

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*  * 本示例向 myDocument 中添加右箭头标注，并且设置该标注的调整值。  * 请注意，尽管形状只有三个调整控点，但是它有四个调整值。第三和第四个调整值都和箭头头部和颈部间的调整句柄相对应。 */ function test() {     let myDocument = Application.Worksheets.Item(1)     let rac = myDocument.Shapes.AddShape(msoShapeRightArrowCallout,10, 10, 250, 190)     let adj = rac.Adjustments     adj.Item(1) = 0.5    //adjusts width of text box     adj.Item(2) = 0.15   //adjusts width of arrow head     adj.Item(3) = 0.8    //adjusts length of arrow head     adj.Item(4) = 0.4    //adjusts width of arrow neck }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Integer** 值，它代表集合中对象的数量。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 返回或设置由 **Index** 参数指定的调整值。**Single** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员属性**

#### **Adjustments.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Adjustments** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET") {       alert("This is an ET Application object.")   }   else {       alert("This is not an ET Application object.")   } }` |

#### **Adjustments.Count**

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Adjustments** 对象的变量。

#### **Adjustments.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Adjustments** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Adjustments.Item**

返回或设置由 **Index** 参数指定的调整值。**Single** 型，可读写。

**语法**

**express.Item**

*express*   一个代表 **Adjustments** 对象的变量。

**说明**

自选图形、连接符和艺术字对象最多可有八个调整。

对于线性调整，调整值 0.0 通常对应于形状的左边缘或上边缘，值 1.0 通常对应于形状的右边缘或下边缘。不过，对于某些形状，调整可以超越边界。对于放射状调整，调整值 1.0 对应于形状的宽度。对于角度的调整，调整值为指定的角度。**Item** 属性只应用于 具有调整的形状。

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                    |
| -------- | ------------- | ------------ | --------------------------- |
| *Index*  | 必选          | **[INT]**    | **Long** 型。调整的索引号。 |

#### **Adjustments.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Adjustments** 对象的变量。

适用环境：web

适用平台：windows/linux