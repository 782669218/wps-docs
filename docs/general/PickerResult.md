#### **PickerResult**



代表解析的或选定的项目数据。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 获取一个 **Application** 对象，该对象代表 **PickerResult** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 获取一个 32 位整数，该整数指示在其中创建了 **PickerResult** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayName**      | 代表 PickerResult 的显示名称。可读/写                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DuplicateResults** | 如果对结果进行解析的结果有多个候选项，则获取 PickerResult 集合。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fields**           | 代表 PickerFields 集合内子项的字段定义。只读                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**               | 检索关联的 **PickerResult** 对象的唯一 Id。只读              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ItemData**         | 获取或设置绑定到数据的不用于显示的项。可读/写                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SIPId**            | 是 Office Communication Server 的标识符。它仅用于人员选取方案。可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SubItems**         | 显示用于显示或不用于显示的 PickerResult 对象的字段数据。它用于在选取器对话框中传递列值。可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**             | 代表 **PickerResult** 对象的类型。可读/写                    |

**成员属性**

#### **PickerResult.Application**

获取一个 **Application** 对象，该对象代表 **PickerResult** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PickerResult** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.DisplayName**

代表 PickerResult 的显示名称。可读/写

**语法**

**express.DisplayName**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.DuplicateResults**

如果对结果进行解析的结果有多个候选项，则获取 PickerResult 集合。只读

**语法**

**express.DuplicateResults**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.Fields**

代表 PickerFields 集合内子项的字段定义。只读

**语法**

**express.Fields**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.Id**

检索关联的 **PickerResult** 对象的唯一 Id。只读

**语法**

**express.Id**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.ItemData**

获取或设置绑定到数据的不用于显示的项。可读/写

**语法**

**express.ItemData**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.SIPId**

是 Office Communication Server 的标识符。它仅用于人员选取方案。可读/写

**语法**

**express.SIPId**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.SubItems**

显示用于显示或不用于显示的 PickerResult 对象的字段数据。它用于在选取器对话框中传递列值。可读/写

**语法**

**express.SubItems**

*express*   一个代表 **PickerResult** 对象的变量。

#### **PickerResult.Type**

代表 **PickerResult** 对象的类型。可读/写

**语法**

**express.Type**

*express*   一个代表 **PickerResult** 对象的变量。

适用环境：web

适用平台：windows/linux