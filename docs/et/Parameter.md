**Parameter**



代表一个在参数查询中使用的参数。

**说明**

**Parameter** 对象是 **Parameters** 集合的成员。

**方法**

|                                                              | 名称         | 说明                   |
| ------------------------------------------------------------ | ------------ | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetParam** | 定义指定查询表的参数。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataType**        | 返回或设置一个 **XlParameterDataType** 值，它代表指定查询参数的数据类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PromptString**    | 返回提示用户在参数查询输入参数值的短语。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshOnChange** | 如果每次更改参数查询的参数值时都要刷新指定的查询表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceRange**     | 返回一个 **Range** 对象，该对象表示包含指定查询参数值的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回一个代表参数类型的 **XlParameterType** 值。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**           | 返回一个代表参数值的 **Variant** 值。                        |

**成员方法**

#### **Parameter.SetParam**

定义指定查询表的参数。

**语法**

**express.SetParam(Type, Value)**

*express*   一个代表 **Parameter** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**        | **说明**                                  |
| -------- | ------------- | ------------------- | ----------------------------------------- |
| *Type*   | 必选          | **XlParameterType** | 指定参数类型的 XlParameterType 常量之一。 |
| *Value*  | 必选          | **Variant**         | 指定参数的值，如 Type 参数的描述所示。    |

**说明**

| XlParameterType 可为以下 XlParameterType 常量之一。          |
| ------------------------------------------------------------ |
| xlConstant。使用 Value 参数指定的值。                        |
| xlPrompt。显示提示用户输入值的对话框。Value 参数指定的是对话框中显示的文字。 |
| xlRange。使用区域左上角单元格的值。Value 参数指定的是一个 Range 对象。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改第一张查询表的 SQL 语句。子句“(city=?)”表示此查询为参数查询，本示例将城市值设置为常量“Oakland”。*/function test(){ 　　　　Set qt = Sheets.Item("sheet1").QueryTables.Item(1) 　　　　qt.Sql = "SELECT * FROM authors  WHERE (city=?)" 　　　　let param1 = qt.Parameters.Add("City Parameter",xlParamTypeVarChar) 　　　　param1.SetParam(xlConstant, "Oakland") 　　　　qt.Refresh()}` |

**成员属性**

#### **Parameter.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Parameter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　if(ActiveWorkbook.Application.Value == "ET") {    　　　　 MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **Parameter.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Parameter** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Parameter.DataType**

返回或设置一个 **XlParameterDataType** 值，它代表指定查询参数的数据类型。

**语法**

**express.DataType**

*express*   一个代表 **Parameter** 对象的变量。

#### **Parameter.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Parameter** 对象的变量。

#### **Parameter.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Parameter** 对象的变量。

#### **Parameter.PromptString**

返回提示用户在参数查询输入参数值的短语。**String** 类型，只读。

**语法**

**express.PromptString**

*express*   一个代表 **Parameter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例修改查询表一的参数提示字符串。*/ function test(){ 　　　　let pa = Worksheets.Item(1).QueryTables.Item(1).Parameters.Item(1) 　　　　pa.SetParam(xlPrompt, "Please " + pa.PromptString) }` |

#### **Parameter.RefreshOnChange**

如果每次更改参数查询的参数值时都要刷新指定的查询表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RefreshOnChange**

*express*   一个代表 **Parameter** 对象的变量。

**说明**

只有当使用 **True**xlRange 类型的参数且引用的参数值位于单个单元格中时，才能将此属性设置为 **xlRange**True。每次更改该单元格的值时，都会刷新。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改 Sheet1 上第一张查询表的 SQL 语句。语句“(ContactTitle=?)”表明该查询是参数查询，标题的值设置为单元格 D4 的值。每次该单元格中的值发生变化时，查询表将会自动刷新。*/ function test(){ 　　　　let objQT = Worksheets.Item("Sheet1").QueryTables.Item(1) 　　　　objQT.CommandText = "Select * From Customers Where (ContactTitle=?)" 　　　　let objParam1 = objQT.Parameters.Add("Contact Title", xlParamTypeVarChar) 　　　　objParam1.RefreshOnChange = true 　　　　objParam1.SetParam(xlRange, Range("D4")) }` |

#### **Parameter.SourceRange**

返回一个 **Range** 对象，该对象表示包含指定查询参数值的单元格。只读。

**语法**

**express.SourceRange**

*express*   一个代表 **Parameter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对作为查询源区域的单元格的值进行更改。*/ function test(){ 　　　　let qt = Sheets.Item("sheet1").QueryTables.Item(1) 　　　　let param1 = qt.Parameters.Item(1) 　　　　let r = param1.SourceRange 　　　　r.Value = "New York" 　　　　qt.Refresh() }` |

#### **Parameter.Type**

返回一个代表参数类型的 **XlParameterType** 值。

**语法**

**express.Type**

*express*   一个代表 **Parameter** 对象的变量。

#### **Parameter.Value**

返回一个代表参数值的 **Variant** 值。

**语法**

**express.Value**

*express*   一个代表 **Parameter** 对象的变量。

适用环境：web

适用平台：windows/linux