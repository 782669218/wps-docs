**AddIn**



代表单个加载宏，不论该加载宏是否已加载。

**说明**

**AddIn** 对象是 **AddIns** 集合的成员。**AddIns** 集合包含 ET 所有可用加载宏的列表（不论这些加载宏是否已安装）。此列表与**“加载宏”**对话框中显示的加载宏列表对应。

示例使用 **AddIns**(*index*)（其中 *index* 是加载宏标题或索引号）可返回单个 **AddIn** 对象。下例安装“分析工具库”加载宏。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.AddIns.Item("analysis toolpak").Installed = true` |

请勿混淆加载宏标题（出现在**“加载宏”**对话框中）与加载宏名称（加载宏的文件名）。必须严格按照**“加载宏”**对话框中的标题书写加载宏标题，但不必匹配大小写。

加载宏索引号代表加载宏在**“加载宏”**对话框内**“可用加载宏”**框中的位置。下例创建一个列表，包含可用加载宏的指定属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let sheet = Application.Worksheets.Item("Sheet1")     sheet.Rows.Item(1).Font.Bold = true     sheet.Range("a1:d1").Value2 = ["Name", "Full Name", "Title", "Installed"]     for(let i = 1; i <= Application.AddIns.Count; i++) {         sheet.Cells(i + 1, 1).Value2 = Application.AddIns.Item(i).Name         sheet.Cells(i + 1, 2).Value2 = Application.AddIns.Item(i).FullName         sheet.Cells(i + 1, 3).Value2 = Application.AddIns.Item(i).Title         sheet.Cells(i + 1, 4).Value2 = Application.AddIns.Item(i).Installed     }     sheet.Range("a1").CurrentRegion.Columns.AutoFit() }` |

**Add** 方法将加载宏添加到可用加载宏列表中，但不安装加载宏。将加载宏的 **Installed** 

属性设为 

True

 可安装加载宏。要安装可用加载宏列表中没有的加载宏，必须先使用 **Add** 

方法，然后设置 **Installed** 

属性。此操作一步即可完成，如下例中所示（注意，**Add** 

方法中应使用加载宏的名称，而不使用标题）。

| 示例代码复制                                             |
| -------------------------------------------------------- |
| `Application.AddIns.Add("generic.xll").Installed = true` |

使用 **Workbooks**

(

index

)（其中 

index

 是加载宏文件名而非标题）可返回对与某一加载宏相对应的工作簿的引用。因为加载宏通常不出现在 **Workbooks** 

集合中，所以必须使用其文件名来指定。此示例将变量 

wb

 设置为 Myaddin.xla 的工作簿。

| 示例代码复制                                         |
| ---------------------------------------------------- |
| `let wb = Application.Workbooks.Item("myaddin.xla")` |

下例将变量 

wb

 设置为“分析工具库”加载宏的工作簿。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let wb = Application.Workbooks.Item(Application.AddIns.Item("analysis toolpak").Name)` |

如果 **Installed** 

属性为 

True

，但调用加载宏中的函数仍旧失败，那么可能并未真正地加载了该加载宏。这是因为 **Addin** 

对象代表了加载宏的存在及安装状态，但并不代表加载宏工作簿中的实际内容。为保证已安装的加载宏被加载，应当打开该加载宏工作簿。下例中，如果加载宏“My Addin”未出现在 **Workbooks** 

集合中，就打开该加载宏。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     try {             // turn off error checking         let wbMyAddin = Application.Workbooks.Item(Application.AddIns.Item("My Addin").Name)         let lastError = Application.Err     }     catch(exception) {         if(lastError != 0) {             // the add-in workbook isn't currently open. Manually open it.             let wbMyAddin = Application.Workbooks.Open(Application.AddIns.Item("My Addin").FullName)         }     } }` |

 

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CLSID**       | 返回一个只读的唯一标识符，或识别对象的 CLSID。**String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**    | 返回对象的名称（以字符串表示），包括其磁盘路径。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Installed**   | 如果安装了此加载宏，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsOpen**      | 如果加载项当前打开，则返回 **True**。只读 **Boolean** 类型   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **progID**      | 返回对象的程序标识符。**String** 型，只读。                  |

**成员属性**

#### **AddIn.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET") {       alert("This is an ET Application object.")   }   else {       alert("This is not an ET Application object.")   } }` |

#### **AddIn.CLSID**

返回一个只读的唯一标识符，或识别对象的 CLSID。**String** 类型。

**语法**

**express.CLSID**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **AddIn.FullName**

返回对象的名称（以字符串表示），包括其磁盘路径。**String** 型，只读。

**语法**

**express.FullName**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示每一个可用加载宏的路径及文件名。*/ function test() {   for(let i = 1; i <= Application.AddIns.Count; i++) {       alert(Application.AddIns.Item(i).FullName)   } }` |

#### **AddIn.Installed**

如果安装了此加载宏，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Installed**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

将此属性设为 **True** 将安装指定加载宏，并调用 Auto_Add 函数。将此属性设为 **False** 则移去指定加载宏，并调用 Auto_Remove 函数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用一个消息框显示 Solver 加载宏的安装状态。*/ function test() {   let a = Application.AddIns.Item("Solver Add-In")   if(a.Installed == true) {       alert("The Solver add-in is installed")   }   else {       alert("The Solver add-in is not installed")   } }` |

#### **AddIn.IsOpen**

如果加载项当前打开，则返回 **True**。只读 **Boolean** 类型

**语法**

**express.IsOpen**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Path**

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.progID**

返回对象的程序标识符。**String** 型，只读。

**语法**

**express.progID**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为第一张工作表中所有 OLE 对象创建程序标识符列表。*/ function test() {   let rw = 0   let o = Application.Worksheets.Item(1).OLEObjects()								   for(let i = 1; i <= o.Count; i++) {       rw = rw + 1       Application.Worksheets.Item(2).Cells.Item(rw, 1).Value2 = o.Item(i).ProgId   } }` |

适用环境：web

适用平台：windows/linux