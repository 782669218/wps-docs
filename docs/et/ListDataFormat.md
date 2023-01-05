**ListDataFormat**



**ListDataFormat** 对象包含 **ListColumn** 对象的所有数据类型属性。这些属性是只读的。

**说明**

使用 ListColumn 对象的 ListDataFormat 属性可返回一个 **ListDataFormat** 对象集合。ListDataFormat 对象的默认属性为表示列表中列的数据类型的 **Type** 属性。这就使得用户在编写代码时无需指定 **Type** 属性。

如下代码示例创建一个与 SharePoint 列表链接的列表。然后检查字段 2 是否为必需（字段 1 为 ID 字段，且只读）。如果该字段是必需的文本字段，在所有现有记录中都编写了相同的数据。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 以下代码示例假设您会将 *strServerName* 和 *strListGuid* 变量替换为有效的服务器名称和列表 GUID。此外，服务器名称后面必须是“/_vti_bin”，否则本示例将不运行。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let strServerName = "http:///_vti_bin" let strListGUID = "{}" let objListObject = Worksheets.Item(1).ListObjects.Add(xlSrcExternal, Array(strServerName, strListGUID), true, xlYes, Range("A1")) let objDataRange = objListObject.ListColumns.Item(2).Range.Offset(1, 0).Resize(objListObject.ListColumns.Item(2).Range.Rows.Count - 2, 1) if(objListObject.ListColumns.Item(2).ListDataFormat.Type == xlListDataTypeText && objListObject.ListColumns.Item(2).ListDataFormat.Required){     objDataRange.Value = "Hello World" } }` |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowFillIn**   | 返回一个 **Boolean** 值，表示用户是否能够为提供值列表的这些列中某一列（而不是限定在值列表范围内）的单元格提供自己的数据。对于未与 SharePoint 网站链接的列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。），返回 **False**。如果未将该列指定为选择或多项选择，也返回 **False**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Choices**       | 返回一个 **String** 值的 **Array**，包含 **DefaultValue** 属性的 **ListLookUp**、**ChoiceMulti** 和 **Choice** 数据类型提供给用户的选项。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DecimalPlaces** | 返回一个 **Long** 值，该值表示要为 **ListColumn** 对象中的数字显示的小数位数。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultValue**  | 返回 **Variant**，代表某列中新的一行的默认数据类型值。如果架构未指定默认值，则返回 **Nothing** 对象。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsPercent**     | 返回一个 **Boolean** 值。只有当 **ListColumn** 对象的数字数据以百分比的格式显示时，才返回 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **lcid**          | 返回一个 **Long** 值，该值表示在架构定义中指定的 **ListColumn** 对象的 LCID。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaxCharacters** | 如果将 **Type** 属性设置为 **xlListDataTypeText** 或 **xlListDataTypeMultiLineText**，则返回 **Long**，该值包含 **ListColumn** 对象中允许的最多字符数。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaxNumber**     | 返回一个 **Variant**，其中包含列表列中该字段允许的最大值。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinNumber**     | 返回一个 **Variant**，其中包含列表列中该字段允许的最小值。该值可以是负的浮点值。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**      | 如果对象以只读方式打开，则返回 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**          | 返回 **XlListDataType** 值，它代表列表列的数据类型。         |

**成员属性**

#### **ListDataFormat.AllowFillIn**

返回一个 **Boolean** 值，表示用户是否能够为提供值列表的这些列中某一列（而不是限定在值列表范围内）的单元格提供自己的数据。对于未与 SharePoint 网站链接的列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。），返回 **False**。如果未将该列指定为选择或多项选择，也返回 **False**。**Boolean** 类型，只读。

**语法**

**express.AllowFillIn**

*express*   一个代表 **ListDataFormat** 对象的变量。

#### **ListDataFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListDataFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListDataFormat.Choices**

返回一个 **String** 值的 **Array**，包含 **DefaultValue** 属性的 **ListLookUp**、**ChoiceMulti** 和 **Choice** 数据类型提供给用户的选项。**Variant** 类型，只读。

**语法**

**express.Choices**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改正在运行 Microsoft SharePoint Foundation 的服务器上的列表来设置这些属性。

**示例**

下例显示了与 SharePoint 列表链接的列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）中第三列的 **Choice** 属性设置。在本例中，假定 **DefaultValue** 属性已设置为 **Choice**、**ChoiceMulti** 或 **ListLookup** 数据类型。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     Debug.Print (objListCol.ListDataFormat.Choices)  }` |

#### **ListDataFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListDataFormat.DecimalPlaces**

返回一个 **Long** 值，该值表示要为 **ListColumn** 对象中的数字显示的小数位数。**Long** 类型，只读。

**语法**

**express.DecimalPlaces**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

如果 **ListDataFormat.Type** 设置与小数位数不符，则返回 0。如果 Microsoft SharePoint Foundation 网站自动确定要显示在 SharePoint 列表中的小数位数，则返回 **xlAutomatic**（-4105 十进制）。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

下例显示了对活动工作簿的 Sheet1 中某个列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）的第三列设置的 **DecimalPlaces** 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)    let GetDecimalPlaces = objListCol.ListDataFormat.DecimalPlaces  }` |

#### **ListDataFormat.DefaultValue**

返回 **Variant**，代表某列中新的一行的默认数据类型值。如果架构未指定默认值，则返回 **Nothing** 对象。**Variant** 类型，只读。

**语法**

**express.DefaultValue**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

该属性只用于链接到 Microsoft SharePoint Foundation 网站的表格。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet1 中表格的第三列设置的 DefaultValue 属性。此代码需要与 SharePoint 网站链接的列表。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)      if(objListCol.ListDataFormat.DefaultValue == null) {         alert("No default value specified.") 	}     else {         alert(objListCol.ListDataFormat.DefaultValue)     }  }` |

#### **ListDataFormat.IsPercent**

返回一个 **Boolean** 值。只有当 **ListColumn** 对象的数字数据以百分比的格式显示时，才返回 **True**。**Boolean** 类型，只读。

**语法**

**express.IsPercent**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

此属性仅用于链接到 Microsoft SharePoint Foundation 网站的列表。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

下例返回了对活动工作簿的 Sheet1 中列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）第三列的 **IsPercent** 属性设置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)      let GetIsPercent = objListCol.ListDataFormat.IsPercent  }` |

#### **ListDataFormat.lcid**

返回一个 **Long** 值，该值表示在架构定义中指定的 **ListColumn** 对象的 LCID。**Long** 类型，只读。

**语法**

**express.lcid**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

在 ET 中，LCID 表示当这是一个 **xlListDataTypeCurrency** 类型时要使用的货币符号。如果没有为列的数据类型设置区域，则返回 0（即 Language Neutral LCID）。

该属性仅用于链接到 Microsoft SharePoint Foundation 网站的表格。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet1 中列表的第三列设置的 lcid 属性。 */ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     alert ("List LCID: " + objListCol.ListDataFormat.lcid)  }` |

#### **ListDataFormat.MaxCharacters**

如果将 **Type** 属性设置为 **xlListDataTypeText** 或 **xlListDataTypeMultiLineText**，则返回 **Long**，该值包含 **ListColumn** 对象中允许的最多字符数。**Long** 类型，只读。

**语法**

**express.MaxCharacters**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

对于 **Type** 属性被设为非文字值的列，则返回 -1。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MaxCharacters 属性。*/ function test(){   let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     Debug.Print (objListCol.ListDataFormat.MaxCharacters) }` |

#### **ListDataFormat.MaxNumber**

返回一个 **Variant**，其中包含列表列中该字段允许的最大值。**Variant** 类型，只读。

**语法**

**express.MaxNumber**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

如果未指定最大值数字，或 **Type** 属性设置使列的最大值不适用，则返回 **Nothing** 对象。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MaxNumber 属性。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     Debug.Print (objListCol.ListDataFormat.MaxNumber) }` |

#### **ListDataFormat.MinNumber**

返回一个 **Variant**，其中包含列表列中该字段允许的最小值。该值可以是负的浮点值。**Variant** 类型，只读。

**语法**

**express.MinNumber**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

如果还没有为该字段指定值，或者 **Type** 属性的设置导致最小值不适用于该列，则该属性将返回 **Nothing** 对象。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MinNumber 属性。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     Debug.Print (objListCol.ListDataFormat.MinNumber) }` |

#### **ListDataFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListDataFormat** 对象的变量。

#### **ListDataFormat.ReadOnly**

如果对象以只读方式打开，则返回 **True**。**Boolean** 类型，只读。

**语法**

**express.ReadOnly**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

此属性仅用于链接到 SharePoint 网站的表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet1 中某个表第三列设置的 ReadOnly 属性。*/ function test(){    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)     Debug.Print (objListCol.ListDataFormat.ReadOnly) }` |

#### **ListDataFormat.Type**

返回 **XlListDataType** 值，它代表列表列的数据类型。

**语法**

**express.Type**

*express*   一个代表 **ListDataFormat** 对象的变量。

**说明**

此属性仅用于链接到 SharePoint 网站的列表。

适用环境：web

适用平台：windows/linux