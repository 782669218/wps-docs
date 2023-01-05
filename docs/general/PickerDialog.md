#### **PickerDialog**



提供用于选择人员或选择数据的对话框用户界面功能。

**说明**

通过 **Application** 对象的 **PickerDialog** 属性获取 **PickerDialog** 对象。

下面的代码设置选取器对话框的属性，然后显示选取器对话框。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `// Configure the Picker Dialog properties. function test() {    	let objPickerDialog = Application.PickerDialog 	objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 	objPickerDialog.Title = "Sample Picker Dialog" 	let objPickerProperties = objPickerDialog.Properties 	let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText) 	let objPickerExistingResults = objPickerDialog.CreatePickerResults() 	let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")   	// Show the Picker Dialog and get the results. 	let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

**方法**

|                                                              | 名称                    | 说明                                                   |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreatePickerResults** | 创建一个空的 **PickerResults** 对象。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Resolve**             | 使用选取器对话框解析标记，并检索结果。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show**                | 使用已指定的数据处理程序和给定的选项显示选取器对话框。 |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 获取一个 **Application** 对象，该对象代表 **PickerDialog** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 获取一个 32 位整数，该整数指示在其中创建了 **PickerDialog** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataHandlerId** | 设置或获取选取器对话框数据处理程序组件的 GUID。可读/写       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Properties**    | 返回 **PickerProperties** 对象，以指定数据处理程序组件的自定义属性。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**         | 设置或返回在选取器对话框中显示的选取器对话框的标题。可读/写  |

**成员方法**

#### **PickerDialog.CreatePickerResults**

创建一个空的 **PickerResults** 对象。

**语法**

**express.CreatePickerResults()**

*express*   一个代表 **PickerDialog** 对象的变量。

**说明**

可以将 PickerResult 添加到返回的对象中，并如同 **PickerDialog** 对象已存在的结果一样，将其指定为 **Show** 方法的第二个参数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的各种属性，并将已存在的 PickerResults 添加到结果中。 function test() {    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerExistingResults = objPickerDialog.CreatePickerResults()    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

#### **PickerDialog.Resolve**

使用选取器对话框解析标记，并检索结果。

**语法**

**express.Resolve(TokenText, duplicateDlgMode)**

*express*   一个代表 **PickerDialog** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**             |
| ------------------ | ------------- | ------------ | -------------------- |
| *TokenText*        | 必选          | **String**   | 要解析的文本字符串。 |
| *duplicateDlgMode* | 必选          | **Integer**  |                      |

**返回值**

PickerResults

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//使用选取器对话框对象解析实体。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)    // Resolve the token by using Picker Dialog and get the results.    let objPickerResults = objPickerDialog.Resolve("johndoe", false) }` |

#### **PickerDialog.Show**

使用已指定的数据处理程序和给定的选项显示选取器对话框。

**语法**

**express.Show(sMultiSelect, ExistingResults)**

*express*   一个代表 **PickerDialog** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**      | **说明**                                                     |
| ----------------- | ------------- | ----------------- | ------------------------------------------------------------ |
| *sMultiSelect*    | 可选          | **Boolean**       | 指定选取器对话框用户界面是否提供多项选择功能。               |
| *ExistingResults* | 必选          | **PickerResults** | 包含选取器对话框用户界面内的现有 PickerResults。这些结果在选定的项控件中显示。 |

**返回值**

PickerResults

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)    let objPickerExistingResults = objPickerDialog.CreatePickerResults()    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")      // Show the Picker Dialog and get the results.    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

**成员属性**

#### **PickerDialog.Application**

获取一个 **Application** 对象，该对象代表 **PickerDialog** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PickerDialog** 对象的变量。

#### **PickerDialog.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PickerDialog** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PickerDialog** 对象的变量。

#### **PickerDialog.DataHandlerId**

设置或获取选取器对话框数据处理程序组件的 GUID。可读/写

**语法**

**express.DataHandlerId**

*express*   一个代表 **PickerDialog** 对象的变量。

**说明**

必须先指定 **DataHandlerID**，之后才能调用选取器对话框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)    let objPickerExistingResults = objPickerDialog.CreatePickerResults()    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")      // Show the Picker Dialog and get the results.    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

#### **PickerDialog.Properties**

返回 **PickerProperties** 对象，以指定数据处理程序组件的自定义属性。只读

**语法**

**express.Properties**

*express*   一个代表 **PickerDialog** 对象的变量。

**说明**

**PickerProperties** 对象的属性将传递给数据处理程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的各种属性并检索结果。 function test() {    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)        // Show the Picker Dialog with no existing result.    let objPickerResults = objPickerDialog.Show(true) }` |

#### **PickerDialog.Title**

设置或返回在选取器对话框中显示的选取器对话框的标题。可读/写

**语法**

**express.Title**

*express*   一个代表 **PickerDialog** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)    let objPickerExistingResults = objPickerDialog.CreatePickerResults()    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")      // Show the Picker Dialog and get the results.    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

适用环境：web

适用平台：windows/linux