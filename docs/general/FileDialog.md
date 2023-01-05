#### **FileDialog**



提供文件对话框，其功能与 WPS Office应用程序中标准的**“打开”**和**“保存”**对话框类似。

**说明**

使用 **FileDialog** 属性返回一个 **FileDialog** 对象。**FileDialog** 属性位于每个单独 Office 应用程序的 **Application** 对象中。该属性使用一个参数 *DialogType* 确定该属性返回的 **FileDialog** 对象类型。**FileDialog** 对象有四种类型：

- **“打开”**对话框 - 允许用户选择一个或多个可以在宿主应用程序中使用 **Execute** 方法打开的文件。
- **“另存为”**对话框 - 允许用户选择一个文件，然后可以使用 **Execute** 方法将当前文件另存为该文件。
- **“文件选取器”**对话框 - 允许用户选择一个或多个文件。用户选择的文件路径将捕获到 **FileDialogSelectedItems** 集合中。
- **“文件夹选取器”**对话框 - 允许用户选择一个路径。用户选择的路径将捕获到 **FileDialogSelectedItems** 集合中。

每个宿主应用程序只能创建一个 **FileDialog** 对象实例。因此，即使创建多个 **FileDialog** 对象，**FileDialog** 对象的很多属性也会保持不变。所以，在显示对话框之前请确保已经针对用途适当地设置了所有属性。

**方法**

|                                                              | 名称        | 说明                                                         |
| ------------------------------------------------------------ | ----------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute** | 在调用 **Show** 方法后立即执行用户的操作。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show**    | 显示文件对话框并返回一个 **Long** 类型的值，指示用户按下的是**“操作”**按钮 (-1) 还是**“取消”**按钮 (0)。在调用 **Show** 方法时，在用户关闭文件对话框之前不会执行其他代码。在**“打开”**和**“另存为”**对话框中，在使用了 **Show** 方法后会立即使用 **Execute** 方法执行用户操作。 |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowMultiSelect** | 如果允许用户从文件对话框中选择多个文件，则为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 获取一个 **Application** 对象，代表 **FileDialog** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ButtonName**       | 设置或获取代表文件对话框中动作按钮上所显示文本的 **String** 类型的值。可读/写。![img]()注释WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 获取一个 32 位整数，指示创建 **FileDialog** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DialogType**       | 返回一个 **MsoFileDialogType** 常量，代表 **FileDialog** 对象被设置为要显示的文件对话框的类型。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FilterIndex**      | 获取或设置一个 **Long** 类型的值，指示文件对话框的默认文件筛选器。默认筛选器决定首次打开文件对话框时显示的文件类型。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Filters**          | 获取一个 **FileDialogFilters** 集合。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InitialFileName**  | 设置或返回一个 **String** 类型的值，代表文件对话框中初始显示的路径或文件名。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InitialView**      | 获取或设置一个 **MsoFileDialogView** 常量，代表文件对话框中文件和文件夹的初始表示形式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**             | 获取与对象关联的文本。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 获取 **FileDialog** 对象的 **Parent** 对象。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SelectedItems**    | 获取一个 **FileDialogSelectedItems** 集合。此集合包含用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选的文件的路径列表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**            | 设置或获取使用 **FileDialog** 对象显示的文件对话框的标题。可读/写。 |

**成员方法**

#### **FileDialog.Execute**

在调用 **Show** 方法后立即执行用户的操作。

**语法**

**express.Execute()**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.Show**

显示文件对话框并返回一个 **Long** 类型的值，指示用户按下的是**“操作”**按钮 (-1) 还是**“取消”**按钮 (0)。在调用 **Show** 方法时，在用户关闭文件对话框之前不会执行其他代码。在**“打开”**和**“另存为”**对话框中，在使用了 **Show** 方法后会立即使用 **Execute** 方法执行用户操作。

**语法**

**express.Show()**

*express*   一个代表 **FileDialog** 对象的变量。

**示例**

**成员属性**

#### **FileDialog.AllowMultiSelect**

如果允许用户从文件对话框中选择多个文件，则为 **True**。可读/写。

**语法**

**express.AllowMultiSelect**

*express*   一个代表 **FileDialog** 对象的变量。

**说明**

此属性对**“文件夹选取器”**对话框或**“另存为”**对话框无效，因为用户永远无法从这些类型的文件对话框中选择多个文件。

#### **FileDialog.Application**

获取一个 **Application** 对象，代表 **FileDialog** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.ButtonName**

设置或获取代表文件对话框中动作按钮上所显示文本的 **String** 类型的值。可读/写。

| ![img]()注释                                           |
| ------------------------------------------------------ |
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

**语法**

**express.ButtonName**

*express*   一个代表 **FileDialog** 对象的变量。

**说明**

默认情况下，此属性设置为该文件对话框类型的标准文本。例如，在**“打开”**对话框中，该属性默认设置为“打开”。此字符串所允许的最大长度为 51 个字符。

#### **FileDialog.Creator**

获取一个 32 位整数，指示创建 **FileDialog** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.DialogType**

返回一个 **MsoFileDialogType** 常量，代表 **FileDialog** 对象被设置为要显示的文件对话框的类型。只读。

**语法**

**express.DialogType**

*express*   一个代表 **FileDialog** 对象的变量。

**示例**

下面的示例将 **FileDialog** 对象视为未知类型，并且如果是**“另存为”**对话框或**“打开”**对话框，则运行 **Execute** 方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(fd){     // If the user presses the action button...     if(fd.Show == -1){         // Use the DialogType property to determine whether to         // use the Execute method.         switch(fd.DialogType){             case Application.Enum.msoFileDialogOpen:             case Application.Enum.msoFileDialogSaveAs:                 fd.Execute()                 break         }          // If the user presses Cancel...     } }` |

#### **FileDialog.FilterIndex**

获取或设置一个 **Long** 类型的值，指示文件对话框的默认文件筛选器。默认筛选器决定首次打开文件对话框时显示的文件类型。可读/写。

**语法**

**express.FilterIndex**

*express*   一个代表 **FileDialog** 对象的变量。

**说明**

如果试图将该属性设置为大于筛选数量的数字，将选定最后一个可用筛选。

**示例**

下面的示例使用 **FileDialog** 对象显示一个**“File Picker”**对话框，并在消息框中显示每个选定的文件。此示例还演示如何添加新文件筛选器并使其成为默认筛选器。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogFilePicker).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for(let i = 1;i <= fdfs.Count;i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files         if(fdf.Extensions.indexOf("xls") > -1) {             alert("Description of filter: " + fdf.Description)         }     }  }` |

#### **FileDialog.Filters**

获取一个 **FileDialogFilters** 集合。只读。

**语法**

**express.Filters**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.InitialFileName**

设置或返回一个 **String** 类型的值，代表文件对话框中初始显示的路径或文件名。可读/写。

**语法**

**express.InitialFileName**

*express*   一个代表 **FileDialog** 对象的变量。

**说明**

在指定文件名时可以使用 **'\*'** 和 **'?'** 通配符，但是指定路径时不能使用这些通配符。**'\*'** 符号代表任意数量的连续字符，而 **'?'** 代表单个字符。例如，**.InitialFileName = "c:\c\*s.txt"** 将返回“charts.txt”和“checkregister.txt”。

如果指定了路径而没有指定文件名，则对话框中将显示文件筛选器所允许的所有文件。

如果指定了位于初始文件夹中的某个文件，则对话框中只显示该文件。

如果指定了初始文件夹中不存在的某个文件名，则对话框中将不包含文件。在 **InitialFileName** 属性中指定的文件类型将覆盖文件筛选器的设置。

如果指定了无效路径，则使用上次使用的路径。如果使用无效路径，则会向用户显示警告消息。

将此属性设置为长度大于 256 个字符的字符串将导致运行时错误。

#### **FileDialog.InitialView**

获取或设置一个 **MsoFileDialogView** 常量，代表文件对话框中文件和文件夹的初始表示形式。可读/写。

**语法**

**express.InitialView**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.Item**

获取与对象关联的文本。只读。

**语法**

**express.Item**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.Parent**

获取 **FileDialog** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FileDialog** 对象的变量。

**说明**

#### **FileDialog.SelectedItems**

获取一个 **FileDialogSelectedItems** 集合。此集合包含用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选的文件的路径列表。只读。

**语法**

**express.SelectedItems**

*express*   一个代表 **FileDialog** 对象的变量。

#### **FileDialog.Title**

设置或获取使用 **FileDialog** 对象显示的文件对话框的标题。可读/写。

**语法**

**express.Title**

*express*   一个代表 **FileDialog** 对象的变量。

适用环境：web

适用平台：windows/linux