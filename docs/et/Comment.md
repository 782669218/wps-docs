**Comment**



代表单元格批注。

**说明**

**Comment** 对象是 **Comments** 集合的成员。

使用 **Comment** 属性可返回 **Comment** 对象。下例更改单元格 E5 中的批注文本。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("E5").Comment.Text("reviewed on " + Date())` |

使用 **Comments**(*index*)（其中 *index* 为批注号）可返回 **Comments** 集合中的单条批注。下例隐藏第一张工作表中的第二条批注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Comments.Item(2).Visible = false` |

使用 **AddComment** 方法可在区域内添加批注。下例在第一张工作表的单元格 E5 中添加批注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()       myComment.Visible = false       myComment.Text("reviewed on " + Date())  } ` |

**方法**

|                                                              | 名称         | 说明                                              |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**   | 删除对象。                                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Next**     | 返回一个 **Comment** 对象，该对象代表下一条批注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Previous** | 返回一个 **Comment** 对象，该对象代表前一条批注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Text**     | 设置批注文本，返回String值。                      |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Author**      | 返回或设置批注的作者。**String** 类型，只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shape**       | 返回一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，它代表连接到指定批注的形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |

**成员方法**

#### **Comment.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Next**

返回一个 **Comment** 对象，该对象代表下一条批注。

**语法**

**express.Next()**

*express*   一个代表 **Comment** 对象的变量。

**说明**

本方法仅对单张工作表有效。对工作表中最后一条批注使用本方法可返回 **Null**（不是下一张工作表的第一条批注）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例隐藏下一条批注*/ /* 请在不带现有批注的新工作簿中进行测试。若要清除工作簿中的所有批注，请在即时窗格中使用 Selection.SpecialCells(xlCellTypeComments).delete*/ //Sets up the comments function test() { 	for(let xNum = 1;xNum <= 10;xNum++) { 		Application.Range("A" + xNum).AddComment() 		Application.Range("A" + xNum).Comment.Text("Comment " + xNum) 	}  	alert("Comments created... A1:A10")  	//Deletes every second comment in the A1:A10 range 	for(let yNum = 1;yNum <= 10;yNum = yNum + 2) { 		Application.Range("A" + yNum).Comment.Next().Shape.Select(true) 		Application.Selection.Delete() 	}  	alert("Deleted every second comment") }` |

#### **Comment.Previous**

返回一个 **Comment** 对象，该对象代表前一条批注。

**语法**

**express.Previous()**

*express*   一个代表 **Comment** 对象的变量。

**说明**

本方法仅对单张工作表有效。对工作表中第一条批注使用本方法可返回 **Null**（不是前一张工作表的最后一条批注）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例隐藏前一条批注*/ /* 请在不带现有批注的新工作簿中进行测试。若要清除工作簿中的所有批注，请在即时窗格中使用Selection.SpecialCells(xlCellTypeComments).Delete*/ //Sets up the comments function test(){ 	//Sets up the comments 	for(let xNum = 1;xNum <= 10;xNum++) { 		Application.Range("A" + xNum).AddComment() 		Application.Range("A" + xNum).Comment.Text("Comment " + xNum) 	}  	alert("Comments created... A1:A10")  	//Deletes every second comment in the A1:A10 range 	for(let yNum = 10;yNum >= 1;yNum = yNum - 2) { 		Range("A" + yNum).Comment.Previous().Shape.Select(true) 		Selection.Delete() 	}  	alert("Deleted every second comment") }` |

#### **Comment.Text**

设置批注文本，返回String值。

**语法**

**express.Text(Text, Start, Overwrite )**

*express*   一个代表 **Comment** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Text*      | 可选          | **Variant**  | 要添加的文本。                                               |
| *Start*     | 可选          | **Variant**  | 所添加文本的起始位置（字符数）。如果省略此参数，则删除批注中的所有现有文字。 |
| *Overwrite* | 可选          | **Variant**  | 如果为 True，则覆盖现有文件。默认值是 False（插入文本）。    |

**成员属性**

#### **Comment.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test(){ 	let myObject = Application.ActiveWorkbook 	if(myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} 	else { 		alert("This is not an ET Application object.") 	} }` |

#### **Comment.Author**

返回或设置批注的作者。**String** 类型，只读。

**语法**

**express.Author**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例删除活动工作表中由 Jean Selva 添加的所有批注*/ function test(){ 	let myComments = Application.ActiveSheet.Comments 	for(let c = 1;c <= myComments.Count;c++) { 		if(myComments.Item(c).Author == "Jean Selva") { 			myComments.Item(c).Delete() 		} 	} }` |

#### **Comment.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Comment** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Comment.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Shape**

返回一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，它代表连接到指定批注的形状。

**语法**

**express.Shape**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例选定活动工作表上的第二条批注*/ /* Ensure that the comments are not hidden. On the Review Tab, choose Comments, Show All Comments*/ Application.ActiveSheet.Comments.Item(2).Shape.Select()` |

#### **Comment.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **Comment** 对象的变量。

适用环境：web

适用平台：windows/linux