**PivotFields**



指定的数据透视表中所有 **PivotField** 对象的集合。

**说明**

下列存取器方法返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- **ColumnFields** 属性
- **DataFields** 属性
- **HiddenFields** 属性
- **PageFields** 属性
- **RowFields** 属性
- **VisibleFields** 属性

使用 **PivotFields** 对象的 **PivotTable** 方法可返回 **PivotFields** 集合。下例列举了 Sheet3 上第一张数据透视表中的字段名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pTab = Worksheets.Item("sheet3").PivotTables(1) for(let i=1;i <= pTab.PivotFields().Count;i++) {     MsgBox(pTab.PivotFields(i).Name) }` |

使用 **PivotFields**( *index*)（其中 *index* 是字段名称或索引号）可返回一个 **PivotField** 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").Orientation = xlRowField` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | table { word-break:break-all; }返回一个 **Long** 值，它代表集合中对象的数量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **PivotFields.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **PivotFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

一个 Object 值，它代表集合包含的某个对象。

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

此示例将工作表 Sheet3 中第一个数据透视表上的 Year 字段设置为行字段。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("sheet3").PivotTables(1).PivotFields().**Item**("year").Orientation = xlRowField` |

**成员属性**

#### **PivotFields.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读

**语法**

**express.Application**

*express*   一个代表 **PivotFields** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotFields.Count**

table { word-break:break-all; }

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **PivotFields** 对象的变量。

#### **PivotFields.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotFields** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotFields.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotFields** 对象的变量。

适用环境：web

适用平台：windows/linux