#### **PickerResults**



**PickerResult** 对象的集合。

**说明**

每个 **PickerResult** 对象

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码显示选取器对话框，获取结果，然后枚举这些结果。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)          // Enumerate the results.     for(let index = 1; index <= objPickerResults.Count-1; index++){         Debug.Print(objPickerResults.Item(index).Id)         Debug.Print(objPickerResults.Item(index).DisplayName)         Debug.Print(objPickerResults.Item(index).Type)         Debug.Print(objPickerResults.Item(index).SIPId)     } }` |

都代表一个解析的或选定的项目数据。

**方法**

|                                                              | 名称    | 说明                                                      |
| ------------------------------------------------------------ | ------- | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 将 **PickerResult** 对象添加到 **PickerResults** 集合中。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **PickerResults** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 检索包含在 **PickerResults** 集合中的 **PickerResult** 对象数的计数。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **PickerResults** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 检索位于指定索引处的 **PickerResult** 对象。只读             |

**成员方法**

#### **PickerResults.Add**

将 **PickerResult** 对象添加到 **PickerResults** 集合中。

**语法**

**express.Add(Id, DisplayName, Type, SIPId, ItemData, SubItems)**

*express*   一个代表 **PickerResults** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Id*          | 必选          | **String**   | 代表 PickerResult 的标识符。                                 |
| *DisplayName* | 必选          | **String**   | 代表 PickerResult 的显示名称。                               |
| *Type*        | 必选          | **String**   | 代表 PickerResult 的类型。                                   |
| *SIPId*       | 可选          | **Boolean**  | 当前不支持。SIPId 是 Office Communication Server 的标识符。它仅用于人员选取方案。 |
| *ItemData*    | 可选          | **Variant**  | 非显示项绑定数据                                             |
| *SubItems*    | 可选          | **Variant**  | 显示用于显示或不用于显示的 PickerResult 的域数据。它用于在选取器对话框中传递列值。 |

**返回值**

PickerResult

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) } ` |

**成员属性**

#### **PickerResults.Application**

获取一个 **Application** 对象，该对象代表 **PickerResults** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PickerResults** 对象的变量。

#### **PickerResults.Count**

检索包含在 **PickerResults** 集合中的 **PickerResult** 对象数的计数。只读

**语法**

**express.Count**

*express*   一个代表 **PickerResults** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码显示选取器对话框用户界面，获取结果，然后枚举这些结果。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)          // Enumerate the results.     for(let index = 1; index <= objPickerResults.Count-1; index++){         Debug.Print(objPickerResults.Item(index).Id)         Debug.Print(objPickerResults.Item(index).DisplayName)         Debug.Print(objPickerResults.Item(index).Type)         Debug.Print(objPickerResults.Item(index).SIPId)     } }` |

#### **PickerResults.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PickerResults** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PickerResults** 对象的变量。

#### **PickerResults.Item**

检索位于指定索引处的 **PickerResult** 对象。只读

**语法**

**express.Item**

*express*   一个代表 **PickerResults** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码显示选取器对话框用户界面，获取结果，然后枚举这些结果。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)          // Enumerate the results.     for(let index = 1; index <= objPickerResults.Count-1; index++){         Debug.Print(objPickerResults.Item(index).Id)         Debug.Print(objPickerResults.Item(index).DisplayName)         Debug.Print(objPickerResults.Item(index).Type)         Debug.Print(objPickerResults.Item(index).SIPId)     } }` |

适用环境：web

适用平台：windows/linux