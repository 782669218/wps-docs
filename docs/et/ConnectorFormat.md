**ConnectorFormat**



包含应用于连接符的属性和方法。

**说明**

连接符是用于连接其他两个形状的线，所连接的位置叫做连接结点。如果重新排列已连接的形状，那么连接符的几何形状将自动调整，以使重新排列的形状仍保持连接。

连接结点通常按下表所示的规则进行编号。

| 形状类型                          | 连接结点标号方案                     |
| --------------------------------- | ------------------------------------ |
| 自选形状、艺术字、图片和 OLE 对象 | 连接结点从顶部开始按逆时针进行编号。 |
| 任意多边形                        | 连接结点为顶点，与顶点编号相对应。   |

使用 **ConnectorFormat** 属性可返回 **ConnectorFormat** 对象。使用 **BeginConnect** 和 **EndConnect** 方法可将连接符的两端连到文档中的其他形状。使用 **RerouteConnections** 方法可自动查找通过连接符连接的两个形状间的最短路径。使用 **Connector** 属性可查看形状是否为连接符。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 请注意，虽然向 **Shapes** 集合添加连接符时，对其设置了大小和位置，但将连接符的起点和终点连接到其他形状时，连接符的大小和位置将会自动调整。因而，如果打算用连接符连接其他形状，那么对其设置的初始大小和位置就没有什么实际意义。同样，用连接符连接其他形状时，将指定要连接的形状上的连接结点，但将连接符连接好之后，用 **RerouteConnections** 方法可能会改变连接符所连接的连接结点，使开始时选定的连接结点变得没有意义。 要算出一个复杂形状上各连接结点的编号，可以打开宏录制器并对形状进行试验操作，然后查看录下的代码；也可以创建一个形状并选中它，然后运行以下示例。这段代码将对每个连接结点进行编号并连接一个连接符。示例代码复制`function test(){ let mainshape = Application.ActiveWindow.Selection.ShapeRange.Item(1)     let bx = mainshape.Left + mainshape.Width + 50     let by = mainshape.Top + mainshape.Height + 50  let mCount = mainshape.ConnectionSiteCount     for(let j = 1;j <= mCount;j++) {          let myConnector = ActiveSheet.Shapes.AddConnector(msoConnectorStraight,                  bx, by, bx + 50, by + 50)             myConnector.ConnectorFormat.EndConnect(mainshape, j)             myConnector.ConnectorFormat.Type = msoConnectorElbow             myConnector.Line.ForeColor.RGB = (255, 0, 0)             let l = myConnector.Left             let t = myConnector.Top                 let myTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal,                  l, t, 36, 14)             myTextbox.Fill.Visible = false             myTextbox.Line.Visible = false             myTextbox.TextFrame.Characters().Text = j      } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向 myDocument 中添加两个矩形并且用曲线连接符连接矩形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let s = myDocument.Shapes let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)     c.ConnectorFormat.BeginConnect(firstRect, 1)     c.ConnectorFormat.EndConnect(secondRect, 1)     c.RerouteConnections() }` |

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BeginConnect**    | 将指定的连接符的起点连接到指定的形状上。如果在连接符的起点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的起点不在所需的连接结点上，本方法将把连接符的起点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **EndConnect** 方法将连接符的终点连接到某一形状上。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BeginDisconnect** | 使指定的连接符的起点与其所连接的形状脱离。本方法并不修改连接符的尺寸和位置：连接符的起点仍保留在原来所连接的连接结点的位置，但与该连接结点之间不再有连接。可用 **EndDisconnect** 方法使连接符的终点与某一形状脱离。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EndConnect**      | 将指定的连接符的终点连接到指定的形状上。如果在连接符的终点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的终点不在所需的连接结点，本方法将把连接符的终点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **BeginConnect** 方法将连接符的起点连接到某一形状上。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EndDisconnect**   | 使指定的连接符的终点与其所连接的形状脱离。本方法并不更改连接符的尺寸和位置：连接符的终点仍保留在原来所连接的连接网站的位置，但与该连接网站之间不再有连接。可以使用 **BeginDisconnect** 方法使连接符的起点与某一形状分离。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginConnected**      | 如果指定的连接符的起点已连接到了某一形状上，则该属性值为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginConnectedShape** | 返回一个 **Shape** 对象，该对象代表指定连接符的起点所连到的形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginConnectionSite** | 返回一个整数，该整数指定连接符起点的连接结点。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndConnected**        | 如果该属性值为 **msoTrue**，则指定的连接符的端点已连接到了某一形状上。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndConnectedShape**   | 返回一个 **Shape** 对象，该对象表示指定连接符的终点所连到的形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndConnectionSite**   | 返回一个整数，该整数指定连接符终点的连接结点。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                | 返回或设置一个 **MsoConnectorType** 值，它代表连接符格式类型。 |

**成员方法**

#### **ConnectorFormat.BeginConnect**

将指定的连接符的起点连接到指定的形状上。如果在连接符的起点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的起点不在所需的连接结点上，本方法将把连接符的起点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **EndConnect** 方法将连接符的终点连接到某一形状上。

**语法**

**express.BeginConnect(ConnectedShape, ConnectionSite)**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ConnectedShape* | 必选          | **Shape**    | 要连接到连接符起点的形状。指定的 Shape 对象必须与连接符在同一个 Shapes 集合中。 |
| *ConnectionSite* | 必选          | **Long**     | 由 ConnectedShape 指定的形状上的连接结点。必须是一个介于 1 和指定形状的 ConnectionSiteCount 属性的返回值之间的整数。如果希望连接符自动查找它所连接的两个形状间的最短路径，请为此参数指定任何一个有效整数，并且在连接符的两端连接到形状之后使用 RerouteConnections 方法。 |

**说明**

将连接符连接到某个对象以后，该连接符的大小和位置将在必要时进行自动调整。

**示例**

本示例向 `myDocument` 中添加了两个矩形，并用曲线连接符将这两个矩形连接起来。请注意，对 **RerouteConnections** 方法的调用使得在 **BeginConnect** 方法和 **EndConnect** 方法中所指定的 *ConnectionSite* 参数值不相关联。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let s = myDocument.Shapes let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) let c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)     c.ConnectorFormat.BeginConnect(firstRect, 1)     c.ConnectorFormat.EndConnect(secondRect, 1)     c.RerouteConnections() }` |

#### **ConnectorFormat.BeginDisconnect**

使指定的连接符的起点与其所连接的形状脱离。本方法并不修改连接符的尺寸和位置：连接符的起点仍保留在原来所连接的连接结点的位置，但与该连接结点之间不再有连接。可用 **EndDisconnect** 方法使连接符的终点与某一形状脱离。

**语法**

**express.BeginDisconnect()**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**示例**

本示例向 `myDocument` 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let s = myDocument.Shapes let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)     c.ConnectorFormat.BeginConnect(firstRect, 1)     c.ConnectorFormat.EndConnect(secondRect, 1)     c.RerouteConnections()     c.ConnectorFormat.BeginDisconnect()     c.ConnectorFormat.EndDisconnect() }` |

#### **ConnectorFormat.EndConnect**

将指定的连接符的终点连接到指定的形状上。如果在连接符的终点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的终点不在所需的连接结点，本方法将把连接符的终点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **BeginConnect** 方法将连接符的起点连接到某一形状上。

**语法**

**express.EndConnect(ConnectedShape, ConnectionSite)**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ConnectedShape* | 必选          | **Shape**    | 要连接到连接符终点的形状。指定的 Shape 对象必须与连接符在同一个 Shapes 集合中。 |
| *ConnectionSite* | 必选          | **Long**     | 是一个介于 1 和指定形状的 ConnectionSiteCount 属性的返回值之间的整数。如果希望连接符自动查找它所连接的两个形状间的最短路径，请为此参数指定任何一个有效整数，并且在连接符的两端连接到形状之后使用 RerouteConnections 方法。 |

**说明**

将连接符连接到某个对象以后，该连接符的大小和位置将在必要时进行自动调整。

**示例**

本示例向 `myDocument` 中添加了两个矩形，并用曲线连接符将这两个矩形连接起来。请注意，对 **RerouteConnections** 方法的调用使得在 **BeginConnect** 方法和 **EndConnect** 方法中所指定的 *ConnectionSite* 参数值不相关联。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let s = myDocument.Shapes let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) let c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)     c.ConnectorFormat.BeginConnect(firstRect, 1)     c.ConnectorFormat.EndConnect(secondRect, 1)     c.RerouteConnections() }` |

#### **ConnectorFormat.EndDisconnect**

使指定的连接符的终点与其所连接的形状脱离。本方法并不更改连接符的尺寸和位置：连接符的终点仍保留在原来所连接的连接网站的位置，但与该连接网站之间不再有连接。可以使用 **BeginDisconnect** 方法使连接符的起点与某一形状分离。

**语法**

**express.EndDisconnect()**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/ function test(){  let myDocument = Application.Worksheets.Item(1) let s = myDocument.Shapes let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)     c.ConnectorFormat.BeginConnect(firstRect, 1)     c.ConnectorFormat.EndConnect(secondRect, 1)     c.RerouteConnections()     c.ConnectorFormat.BeginDisconnect()     c.ConnectorFormat.EndDisconnect() }` |

**成员属性**

#### **ConnectorFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     MsgBox("This is an ET Application object.") } else {     MsgBox("This is not an ET Application object.") } }` |

#### **ConnectorFormat.BeginConnected**

如果指定的连接符的起点已连接到了某一形状上，则该属性值为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.BeginConnected**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue**：指定连接符的起点与形状相连。             |

**示例**

如果 `myDocument` 上的第三个形状是连接符，且它的起点已连接到了某一形状上，本示例将把连接结点的编号存储到变量 `oldBeginConnSite` 中，把对所连接的形状的引用存储到对象变量 `oldBeginConnShape` 中，然后断开连接符的起点与形状的连接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes.Item(3)     if(myShape.Connector) {         let myFormat = myShape.ConnectorFormat             if(myFormat.BeginConnected) {                 let oldBeginConnSite = myFormat.BeginConnectionSite                 let oldBeginConnShape = myFormat.BeginConnectedShape                 myFormat.BeginDisconnect()             }     } }` |

#### **ConnectorFormat.BeginConnectedShape**

返回一个 **Shape** 对象，该对象代表指定连接符的起点所连到的形状。只读。

**语法**

**express.BeginConnectedShape**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

如果指定的连接符的起点并未连接到任何形状上，那么本属性将产生错误。

**示例**

本示例假定在 `myDocument` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 `myDocument` 添加一个矩形和一个连接符。新添加的连接符的起点将连接到“Conn1To2”的起点所连接的同一连接结点上，而新添加的连接符的终点则连接到新添加的矩形的第一个连接结点上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes     let r3 = myShape.AddShape(msoShapeRectangle, 450, 190, 200, 100)     myShape.AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = "Conn1To3"     let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat         let beginConnSite1 = connFormat1.BeginConnectionSite         let beginConnShape1 = connFormat1.BeginConnectedShape     let connFormat2 = myShape.Item("Conn1To3").ConnectorFormat         connFormat2.BeginConnect(beginConnShape1, beginConnSite1)         connFormat2.EndConnect(r3, 1) }` |

#### **ConnectorFormat.BeginConnectionSite**

返回一个整数，该整数指定连接符起点的连接结点。**Long** 类型，只读。

**语法**

**express.BeginConnectionSite**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

如果指定的连接符的起点并未连接到任何形状上，那么本属性将产生错误。

**示例**

本示例假定在 `myDocument` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 `myDocument` 添加一个矩形和一个连接符。新添加的连接符的起点将连接到“Conn1To2”的起点所连接的同一连接结点上，而新添加的连接符的终点则连接到新添加的矩形的第一个连接结点上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes     let r3 = myShape.AddShape(msoShapeRectangle, 450, 190, 200, 100)     myShape.AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = "Conn1To3"     let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat         let beginConnSite1 = connFormat1.BeginConnectionSite         let beginConnShape1 = connFormat1.BeginConnectedShape     let connFormat2 = myShape.Item("Conn1To3").ConnectorFormat         connFormat2.BeginConnect(beginConnShape1, beginConnSite1)         connFormat2.EndConnect(r3, 1) }` |

#### **ConnectorFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ConnectorFormat.EndConnected**

如果该属性值为 **msoTrue**，则指定的连接符的端点已连接到了某一形状上。**MsoTriState** 类型，只读。

**语法**

**express.EndConnected**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue** 不应用于该属性。                         |
| **msoFalse** 指定连接符的端点未连接到某一形状上。     |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于该属性。                |
| **msoTrue** 指定连接符的端点连接到形状上。            |

**示例**

如果 `myDocument` 上的第三个形状是连接符，并且它的终点已连接到了某一形状上，则本示例将把连接结点的编号存储到变量 `oldEndConnSite` 中，把对所连接的形状的引用存储到对象变量 `oldEndConnShape` 中，然后断开连接符的终点与形状的连接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes.Item(3)     if(myShape.Connector) {         let myFormat = myShape.ConnectorFormat             if(myFormat.EndConnected) {                 let oldEndConnSite = myFormat.EndConnectionSite                 let oldEndConnShape = myFormat.EndConnectedShape                 myFormat.EndDisconnect()             }     } }` |

#### **ConnectorFormat.EndConnectedShape**

返回一个 **Shape** 对象，该对象表示指定连接符的终点所连到的形状。只读。

**语法**

**express.EndConnectedShape**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

如果指定的连接符的终点没有连接到一个形状上，则该属性产生一个错误。

**示例**

本示例假定在 `myDocument` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 `myDocument` 添加一个矩形和一个连接符。新添加的连接符的终点将连接到“Conn1To2”的终点所连接的同一连接结点上，而新添加的连接符的起点则连接到新添加的矩形的第一个连接结点上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes     let r3 = myShape.AddShape(msoShapeRectangle, 100, 420, 200, 100)     let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat         let endConnSite1 = connFormat1.EndConnectionSite         let endConnShape1 = connFormat1.EndConnectedShape     let connFormat2 = myShape.AddConnector(msoConnectorCurve,              0, 0, 10, 10).ConnectorFormat         connFormat2.BeginConnect(r3, 1)         connFormat2.EndConnect(endConnShape1, endConnSite1) }` |

#### **ConnectorFormat.EndConnectionSite**

返回一个整数，该整数指定连接符终点的连接结点。**Long** 类型，只读。

**语法**

**express.EndConnectionSite**

*express*   一个代表 **ConnectorFormat** 对象的变量。

**说明**

如果指定的连接符的终点没有连接到一个形状上，则该属性产生一个错误。

**示例**

本示例假定在 `myDocument` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 `myDocument` 添加一个矩形和一个连接符。新添加的连接符的终点将连接到“Conn1To2”的终点所连接的同一连接结点上，而新添加的连接符的起点则连接到新添加的矩形的第一个连接结点上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myDocument = Application.Worksheets.Item(1) let myShape = myDocument.Shapes     let r3 = myShape.AddShape(msoShapeRectangle, 100, 420, 200, 100)     let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat         let endConnSite1 = connFormat1.EndConnectionSite         let endConnShape1 = connFormat1.EndConnectedShape     let connFormat2 = myShape.AddConnector(msoConnectorCurve,              0, 0, 10, 10).ConnectorFormat         connFormat2.BeginConnect(r3, 1)         connFormat2.EndConnect(endConnShape1, endConnSite1) }` |

#### **ConnectorFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ConnectorFormat** 对象的变量。

#### **ConnectorFormat.Type**

返回或设置一个 **MsoConnectorType** 值，它代表连接符格式类型。

**语法**

**express.Type**

*express*   一个代表 **ConnectorFormat** 对象的变量。

适用环境：web

适用平台：windows/linux