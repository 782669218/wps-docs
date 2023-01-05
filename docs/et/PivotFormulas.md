**PivotFormulas**



代表数据透视表的公式的集合。每个公式都由一个 **PivotFormula** 对象代表。

**说明**

本对象及其相关属性和方法对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 **PivotFormulas** 属性可返回 **PivotFormulas** 集合。下例为活动工作表上的第一张数据透视表创建一个数据透视表公式列表。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let r = 1 for(let i=1;i <= ActiveSheet.PivotTables(1).**PivotFormulas**.Count;i++) {     Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula     r++ }` |

**方法**

|                                                              | 名称        | 说明                                                         |
| ------------------------------------------------------------ | ----------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**     | 新建数据透视表公式。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Creator** | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**    | 从集合中返回一个对象。                                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | table { word-break:break-all; }返回一个 **Long** 值，它代表集合中对象的数量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | table { word-break:break-all; }返回指定对象的父对象。只读。  |

**成员方法**

#### **PivotFormulas.Add**

新建数据透视表公式。

**语法**

**express.Add(Formula, UseStandardFormula)**

*express*   一个代表 **PivotFormulas** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**             |
| -------------------- | ------------- | ------------ | -------------------- |
| *Formula*            | 必选          | **String**   | 新的数据透视表公式。 |
| *UseStandardFormula* | 可选          | **Variant**  | 标准数据透视表公式。 |

**返回值**

一个 PivotFormula 对象，它代表新的数据透视表公式。

**示例**

此示例在第一张工作表上为第一个数据透视表创建一个新的数据透视表公式。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables(1).PivotFormulas.Add("Year['1998'] Apples = (Year['1997'] Apples) * 2")` |

#### **PivotFormulas.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator()**

*express*   一个代表 **PivotFormulas** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotFormulas.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **PivotFormulas** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 PivotFormula 对象。

**示例**

本示例显示第一张工作表中第一个数据透视表的第一个公式。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `MsgBox(Worksheets.Item(1).PivotTables(1).PivotFormulas.**Item**(1).Formula)` |

**成员属性**

#### **PivotFormulas.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotFormulas** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotFormulas.Count**

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **PivotFormulas** 对象的变量。

#### **PivotFormulas.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotFormulas** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotFormulas.Parent**

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotFormulas** 对象的变量。

适用环境：web

适用平台：windows/linux