**Error**



代表区域的电子表格错误。

**说明**

使用 **Errors** 对象的 **Item** 属性可返回 **Error** 对象。

返回 **Error** 对象后，可结合使用 **Value** 和 **Errors** 属性检查某个特定错误检查选项是否已启用。

| ![img]()注释                                              |
| --------------------------------------------------------- |
| 不要将 **Error** 对象与 Visual Basic 的错误处理功能混淆。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例中，在引用空单元格的单元格 A1 中创建一个公式，然后使用 Item(index)（其中 index 用于标识错误类型）显示描述错误情况的消息。*/ function test(){  let rngFormula = Application.Range("A1")      //Place a formula referencing empty cells.     Range("A1").Formula = "=A2+A3"     Application.ErrorCheckingOptions.EmptyCellReferences = true      //Perform check to see if EmptyCellReferences check is on.     if(rngFormula.Errors.Item(xlEmptyCellReferences).Value == true) {         alert("The empty cell references error checking feature is enabled.")     }     else {         alert("The empty cell references error checking feature is not on.")     } }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Ignore**      | 允许用户设置或返回某一区域的错误检查选项的状态。如果该值为 **False**，则某一区域的错误检查选项可用。如果该值为 **True**，则禁用某一区域的错误检查选项。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。 |

**成员属性**

#### **Error.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Error** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **Error.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Error** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Error.Ignore**

允许用户设置或返回某一区域的错误检查选项的状态。如果该值为 **False**，则某一区域的错误检查选项可用。如果该值为 **True**，则禁用某一区域的错误检查选项。**Boolean** 类型，可读写。

**语法**

**express.Ignore**

*express*   一个代表 **Error** 对象的变量。

**说明**

引用 **ErrorCheckingOptions** 对象可查看与错误检查选项相关联的索引值列表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为检查空单元格引用而禁用单元格 A1 中的忽略标记。*/ function test(){  Range("A1").Select()      //Determine if empty cell references error checking is on, if not turn it on.     if(Application.Range("A1").Errors.Item(xlEmptyCellReferences).Ignore == true) {         Application.Range("A1").Errors.Item(xlEmptyCellReferences).Ignore = false         alert("Empty cell references error checking has been enabled for cell A1.")     }     else {         alert("Empty cell references error checking is already enabled for cell A1.")     } }` |

#### **Error.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Error** 对象的变量。

#### **Error.Value**

返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。

**语法**

**express.Value**

*express*   一个代表 **Error** 对象的变量。

适用环境：web

适用平台：windows/linux