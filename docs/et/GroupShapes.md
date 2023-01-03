**GroupShapes**



代表一组形状中的单个形状。

**说明**

每个形状都由一个 **Shape** 对象代表。将 **Item** 方法与此对象一起使用，您可以在不取消分组的情况下处理组合的各个形状。

使用 **GroupItems** 属性可返回 **GroupShapes** 集合。使用 **GroupItems**(*index* )（其中 *index* 是分组的形状中单个形状的编号）可从 **GroupShapes** 集合中返回单个形状。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向 myDocument 添加三个三角形，将它们分成一组，设置整个组的颜色，然后只更改第二个三角形的颜色。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let shape1 = myDocument.Shapes     shape1.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"     shape1.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"     shape1.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"     let shape2 = shape1.Range(["shpOne", "shpTwo", "shpThree"]).Group()         shape2.Fill.PresetTextured(msoTextureBlueTissuePaper)         shape2.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble) }` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回一个 **ShapeRange** 对象，它代表 **Shapes** 集合中形状的子集。 |

**成员方法**

#### **GroupShapes.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **GroupShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Shape 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对某个形状区域中第二个形状的 OnAction 属性进行设置。如果变量 sr 不代表 ShapeRange 对象，则本示例无效。*/ function test(){ let sr = Application.Worksheets.Item(1).Shapes sr.Item(2).OnAction = "ShapeAction" }` |

**成员属性**

#### **GroupShapes.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **GroupShapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **GroupShapes.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **GroupShapes** 对象的变量。

#### **GroupShapes.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **GroupShapes** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **GroupShapes.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **GroupShapes** 对象的变量。

#### **GroupShapes.Range**

返回一个 **ShapeRange** 对象，它代表 **Shapes** 集合中形状的子集。

**语法**

**express.Range**

*express*   一个代表 **GroupShapes** 对象的变量。

**说明**

虽然使用 **Range** 属性可返回任意数量的形状，但如果要返回集合的单个成员，用 **Item** 方法更加简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 myDocument 中第一个形状和第三个形状的填充图案。*/ function test(){ let myDocument = Application.Worksheets.Item(1) myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*若要为 Index 指定一个整数或字符串数组，可以使用 Array 函数。例如，以下指令返回用名称指定的两个形状。*/ function test(){ let arShapes = ["Oval 4", "Rectangle 5"] let objRange = Application.ActiveSheet.Shapes.Range(arShapes) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在 ET 中，不能用此属性返回包含工作表上的所有 Shape 对象的 ShapeRange 对象。如果要达到该目的，可用下列代码： */ function test(){ Application.Worksheets.Item(1).Shapes.SelectAll	    // select all shapes let sr = Selection.ShapeRange	  	    // create ShapeRange }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 myDocument 中形状“Oval 4”和“Rectangle 5”的填充图案。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let arShapes = ["Oval 4", "Rectangle 5"] let objRange = myDocument.Shapes.Range(arShapes) objRange.Fill.Patterned(msoPatternHorizontalBrick) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 myDocument 中第一个形状的填充图案。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let myRange = myDocument.Shapes.Range(1) myRange.Fill.Patterned(msoPatternHorizontalBrick) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例创建一个数组，其中包含 myDocument 中所有的自选图形，并用该数组定义一个形状区域，然后在该区域中水平分布所有的形状。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let shapes2 = myDocument.Shapes     numShapes = shapes2.Count     if(numShapes > 1) {         let numAutoShapes = 0         let autoShpArray = []         for(let i = 1; i <= numShapes; i++) {             if(shapes2.Item(i).Type == msoAutoShape) {                 autoShpArray[numAutoShapes] = shapes2.Item(i).Name                 numAutoShapes++             }         }         if(numAutoShapes > 1) {             let asRange = shapes2.Range(autoShpArray)             asRange.Distribute(msoDistributeHorizontally, false)         }     } }` |

适用环境：web

适用平台：windows/linux