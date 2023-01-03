#### **PickerProperty**



代表一个对象，以便传递自定义属性。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {     // Configure the Picker Dialog properties.     let objPickerDialog = Application.PickerDialog     objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"     objPickerDialog.Title = "Sample Picker Dialog"     let objPickerProperties = objPickerDialog.Properties     let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)     let objPickerExistingResults = objPickerDialog.CreatePickerResults()     let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")           // Show the Picker Dialog and get the results.     let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **PickerProperty** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **PickerProperty** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**          | 检索关联的 **PickerProperty** 对象的唯一 Id。只读            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 检索选取器属性的类型。只读                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 检索选取器属性的值。只读                                     |

**成员属性**

#### **PickerProperty.Application**

获取一个 **Application** 对象，该对象代表 **PickerProperty** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PickerProperty** 对象的变量。

#### **PickerProperty.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PickerProperty** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PickerProperty** 对象的变量。

#### **PickerProperty.Id**

检索关联的 **PickerProperty** 对象的唯一 Id。只读

**语法**

**express.Id**

*express*   一个代表 **PickerProperty** 对象的变量。

#### **PickerProperty.Type**

检索选取器属性的类型。只读

**语法**

**express.Type**

*express*   一个代表 **PickerProperty** 对象的变量。

#### **PickerProperty.Value**

检索选取器属性的值。只读

**语法**

**express.Value**

*express*   一个代表 **PickerProperty** 对象的变量。

适用环境：web

适用平台：windows/linux