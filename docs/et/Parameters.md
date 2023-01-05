**Parameters**



指定的查询表的 **Parameter** 对象的集合。

**说明**

每一个 **Parameter** 对象都代表一个查询参数。每一个查询表都包含一个 **Parameters** 集合，但集合是空的，除非查询表使用的是参数查询。

您不能在 URL 连接查询表上使用 **Add** 方法。对于 URL 连接查询表，ET 会基于 **Connection** 和 **PostText** 属性创建参数。

**方法**

|                                                              | 名称       | 说明                   |
| ------------------------------------------------------------ | ---------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 创建新查询参数。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | table { word-break:break-all; }返回指定对象的父对象。只读。  |

**成员方法**

#### **Parameters.Add**

创建新查询参数。

**语法**

**express.Add(Name, iDataType)**

*express*   一个代表 **Parameters** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*      | 必选          | **String**   | 指定参数的名称。该参数名称必须与 SQL 语句中的参数子句相符。  |
| *iDataType* | 必选          | **Variant**  | 参数的数据类型。可以为任何 XlParameterDataType 常量。这些值与 ODBC 数据类型相对应，它们指明 ODBC 驱动程序要接收的值类型。ET 和 ODBC 驱动程序管理器将对 ET 提供的参数值进行强制转换，使之成为 ODBC 驱动程序可接受的正确数据类型。 |

**返回值**

一个代表新查询参数的 Parameter 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改第一张查询表的 SQL 语句。子句“(city=?)”表明此查询为参数查询，城市值被设置为常量“Oakland”。*/ function test(){ 　　　　let qt = Sheets.Item("sheet1").QueryTables.Item(1) 　　　　qt.Sql = "SELECT * FROM authors  WHERE (city=?)" 　　　　let param1 = qt.Parameters.Add("City Parameter", xlParamTypeVarChar) 　　　　param1.SetParam (xlConstant, "Oakland") 　　　　qt.Refresh() }` |

#### **Parameters.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Parameters** 对象的变量。

#### **Parameters.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Parameters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Parameter 对象。

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例修改参数提示字符串。*/ function test(){ 　　　　let pa = Worksheets.Item(1).QueryTables.Item(1).Parameters.Item(1) 　　　　pa.SetParam(xlPrompt, "Please " + pa.PromptString) }` |

**成员属性**

#### **Parameters.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Parameters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　if(ActiveWorkbook.Application.Value == "ET") { 　　　　    MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **Parameters.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Parameters** 对象的变量。

#### **Parameters.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Parameters** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Parameters.Parent**

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Parameters** 对象的变量。

适用环境：web

适用平台：windows/linux