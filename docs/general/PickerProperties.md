#### **PickerProperties**



**PickerProperty** 对象的集合。

**说明**

每个 **PickerProperty** 对象都是一个名称 (ID)/值对，用于将选项值传递给 PickerDialog 对象。可以通过 **PickerDialog** 对象的 **Properties** 属性获取 **PickerProperties** 集合对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置选取器对话框的属性，然后显示选取器对话框。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)    let objPickerExistingResults = objPickerDialog.CreatePickerResults()    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")         // Show the Picker Dialog and get the results.    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult) }` |

**方法**

|                                                              | 名称       | 说明 |
| ------------------------------------------------------------ | ---------- | ---- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    |      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Remove** |      |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **PickerProperties** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 检索包含在 **PickerProperties** 集合中的 **PickerProperty** 对象数的计数。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **PickerProperties** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 检索位于指定索引处的 **PickerProperty** 对象。只读           |

**成员方法**

#### **PickerProperties.Add**

**语法**

**express.Add(Id, Value, Type)**

*express*   一个代表 **PickerProperties** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**       | **说明**       |
| -------- | ------------- | ------------------ | -------------- |
| *Id*     | 必选          | **String**         | 属性的项名称。 |
| *Value*  | 必选          | **Boolean**        | 属性的值。     |
| *Type*   | 必选          | **MsoPickerField** | 属性的类型。   |

**返回值**

PickerProperty

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码设置 PickerDialog 对象的各种属性。 function test() {    // Configure the Picker Dialog properties.    let objPickerDialog = Application.PickerDialog    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"    objPickerDialog.Title = "Sample Picker Dialog"    let objPickerProperties = objPickerDialog.Properties    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldtypeText) }` |

#### **PickerProperties.Remove**

**语法**

**express.Remove(Id)**

*express*   一个代表 **PickerProperties** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *Id*     | 必选          | **String**   | 要删除的 PickerProperty 的标识符。 |

**成员属性**

#### **PickerProperties.Application**

获取一个 **Application** 对象，该对象代表 **PickerProperties** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PickerProperties** 对象的变量。

#### **PickerProperties.Count**

检索包含在 **PickerProperties** 集合中的 **PickerProperty** 对象数的计数。只读

**语法**

**express.Count**

*express*   一个代表 **PickerProperties** 对象的变量。

#### **PickerProperties.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PickerProperties** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PickerProperties** 对象的变量。

#### **PickerProperties.Item**

检索位于指定索引处的 **PickerProperty** 对象。只读

**语法**

**express.Item**

*express*   一个代表 **PickerProperties** 对象的变量。

适用环境：web

适用平台：windows/linux