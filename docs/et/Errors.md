**Errors**



------

代表区域内的各种电子表格错误。

**说明**

使用 **Range** 对象的 **Errors** 属性可返回 **Errors** 对象。

返回 **Errors** 对象后，可使用 **Error** 对象的 **Value** 属性检查特定的错误检查条件。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将一个数字作为文本放在单元格 A1 中，然后当单元格 A1 的值包含文本格式的数字时通知用户。*/ function test(){  //Place a number written as text in cell A1.     Range("A1").Formula = "'1"      if(Range("A1").Errors.Item(xlNumberAsText).Value == true) {         alert("Cell A1 has a number as text.")     }     else {         alert("Cell A1 is a number.")     }  }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 返回 **Error** 对象的单个成员。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员属性**

#### **Errors.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Errors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **Errors.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Errors** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Errors.Item**

返回 **Error** 对象的单个成员。

**语法**

**express.Item**

*express*   一个代表 **Errors** 对象的变量。

**说明**

*Index* 可以是以下常量之一。

| **xlEvaluateToError**：单元格计算为错误值。                  |
| ------------------------------------------------------------ |
| **xlTextDate**：单元格包含用 2 位数表示年份的文本日期。      |
| **xlNumberAsText**：单元格包含以文本形式存储的数字。         |
| **xlInconsistentFormula**：单元格包含一个区域中不一致的公式。 |
| **xlOmittedCells**：单元格包含一个忽略了区域中某个单元格的公式。 |
| **xlUnlockedFormulaCells**：取消锁定的单元格包含一个公式。   |
| **xlEmptyCellReferences**：单元格包含一个引用空单元格的公式。 |

#### **Errors.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Errors** 对象的变量。

适用环境：web

适用平台：windows/linux