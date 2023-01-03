#### **FileDialogSelectedItems**



**String** 值的集合，这些值对应于通过 **FileDialog** 对象显示的文件对话框中用户所选文件或文件夹路径。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     //Declare a variable as a FileDialog object.      //Create a FileDialog object as a File Picker dialog box.     let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)      //Declare a variable to contain the path     //of each selected item. Even though the path is aString,     //the variable must be a Variant because For Each...Next     //routines only work with Variants and Objects.      //Use a With...End With block to reference the FileDialog object.                       //Allow the selection of multiple file.          fd.AllowMultiSelect = true           //Use the Show method to display the File Picker dialog box and return the user's action.         //The user pressed the button.         if(fd.Show() == -1) {              //Step through each string in the FileDialogSelectedItems collection             for(let i = 1;i <= fd.SelectedItems.Count;i++) {                 let vrtSelectedItem = fd.SelectedItems.Item(i)                 //vrtSelectedItem is aString that contains the path of each selected item.                 //You can use any file I/O functions that you want to work with this path.                 //This example displays the path in a message box.                 alert("Selected item's path: " + vrtSelectedItem)              }         }         //The user pressed Cancel.         else { }               //Set the object variable to Nothing.     fd = null  }` |

使用 **FileDialog** 对象的 **SelectedItems** 属性可返回一个 **FileDialogSelectedItems** 集合。下面的示例显示一个**“File Picker”**对话框，并在消息框中显示每个选定的文件。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 获取一个 **String** 类型的值，对应于用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选择的文件之一的路径。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **FileDialogSelectedItems** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **FileDialogSelectedItem**s 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **FileDialogSelectedItems** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **FileDialogSelectedItems** 对象的 **Parent** 对象。只读。 |

**成员方法**

#### **FileDialogSelectedItems.Item**

获取一个 **String** 类型的值，对应于用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选择的文件之一的路径。

**语法**

**express.Item()**

*express*   一个代表 **FileDialogSelectedItems** 对象的变量。

**成员属性**

#### **FileDialogSelectedItems.Application**

获取一个 **Application** 对象，代表 **FileDialogSelectedItems** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **FileDialogSelectedItems** 对象的变量。

#### **FileDialogSelectedItems.Count**

获取一个 **Long** 类型的值，指示 **FileDialogSelectedItem**s 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **FileDialogSelectedItems** 对象的变量。

**示例**

下面的示例显示一个“File Picker”对话框，并在消息框中显示每个选定的文件。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialog object.      //Create a FileDialog object as a File Picker dialog box.     let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)      //Declare a variable to contain the path     //of each selected item. Even though the path is aString,     //the variable must be a Variant because For Each...Next     //routines only work with Variants and Objects.      //Use a With...End With block to reference the FileDialog object.                      //Allow the selection of multiple file.         fd.AllowMultiSelect = true           //Use the Show method to display the File Picker dialog box and return the user's action.         //The user pressed the button.         if(fd.Show() == -1) {              //Step through each string in the FileDialogSelectedItems collection             for(let cnt = 1;cnt <= fd.SelectedItems.Count;cnt++) {                 let vrtSelectedItem = fd.SelectedItems                 //vrtSelectedItem is aString that contains the path of each selected item.                 //You can use any file I/O functions that you want to work with this path.                 //This example displays the path in a message box.                 alert("Selected item's path: " + vrtSelectedItem.Item(cnt))              }         }         //The user pressed Cancel.         else { }       //Set the object variable to Nothing.     fd = null  }` |

#### **FileDialogSelectedItems.Creator**

获取一个 32 位整数，指示创建 **FileDialogSelectedItems** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **FileDialogSelectedItems** 对象的变量。

#### **FileDialogSelectedItems.Parent**

获取 **FileDialogSelectedItems** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FileDialogSelectedItems** 对象的变量。

**类型**

Object

适用环境：web

适用平台：windows/linux