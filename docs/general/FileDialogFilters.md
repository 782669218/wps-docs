#### **FileDialogFilters**



表示文件对话框中可选文件类型的 **FileDialogFilter** 对象的集合，该对话框由 **FileDialog** 对象显示。

**说明**

使用 **FileDialog** 对象的 **Filters** 属性可返回一个 **FileDialogFilters** 集合。

使用 Add 方法可向 FileDialogFilters 集合中添加 FileDialogFilter 对象。下面的示例使用 Clear 方法清除集合，然后向集合中添加筛选器。Clear 方法将集合完全清空；但是，如果在清除后没有向集合中添加任何筛选器，将自动添加“所有文件(*.*)”筛选器。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function Main() {      //Declare a variable as a FileDialog object.      //Create a FileDialog object as a File Picker dialog box.     let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)      //Declare a variable to contain the path     //of each selected item. Even though the path is aString,     //the variable must be a Variant because For Each...Next     //routines only work with Variants and Objects.      //Use a With...End With block to reference the FileDialog object.          //Change the contents of the Files of Type list.         //Empty the list by clearing the FileDialogFilters collection.         fd.Filters.Clear()          //Add a filter that includes all files.         fd.Filters.Add("All files", "*.*")          //Add a filter that includes GIF and JPEG images and make it the first item in the list.         fd.Filters.Add("Images", "*.gif; *.jpg; *.jpeg", 1)          //Use the Show method to display the File Picker dialog box and return the user's action.         //The user pressed the button.         if(fd.Show() == -1) {              //Step through eachString in the FileDialogSelectedItems collection.             for(let i = 1;i <= fd.SelectedItems.Count;i++){                 let vrtSelectedItem = fd.SelectedItems.Item(i)                 //vrtSelectedItem is aString that contains the path of each selected item.                 //You can use any file I/O functions that you want to work with this path.                 //This example displays the path in a message box.                 MsgBox("Path name: " + vrtSelectedItem)              }         }         //The user pressed Cancel.         else { }      //Set the object variable to Nothing.     fd = null  }` |

在更改 **FileDialogFilters** 集合时，请记住每个应用程序只能创建单个 **FileDialog** 对象的实例。这意味着在调用一个新对话框类型的 **FileDialog** 方法时，**FileDialogFilters** 集合将重置为其默认筛选器。

下面的示例循环访问**“SaveAs”**对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for(let i = 1;i <= fdfs.Count;i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files         if(fdf.Extensions.indexOf("xls") > -1) {             alert("Description of filter: " + fdf.Description)         }     }  }` |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 当应用于“另存为”**FileDiaog** 对象时，如果 **Filters** 属性与 **Clear**、**Add** 或 **Delete** 方法结合使用，则会发生运行时错误。例如，`Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters.Clear()`会引起运行时错误。 |

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 在**“文件”**对话框的**“文件类型”**下拉列表框的筛选器列表中添加一个新的文件筛选器。返回一个代表新添的文件筛选器的 **FileDialogFilter** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear**  | 删除当前在文件对话框中应用的所有筛选器。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除文件对话框筛选。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 获取属于指定的 **FileDialogFilters** 集合的成员的 **FileDialogFilter** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **FileDialogFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **FileDialogFilters** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **FileDialogFilters** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **FileDialogFilters** 对象的 **Parent** 对象。只读。    |

**成员方法**

#### **FileDialogFilters.Add**

在**“文件”**对话框的**“文件类型”**下拉列表框的筛选器列表中添加一个新的文件筛选器。返回一个代表新添的文件筛选器的 **FileDialogFilter** 对象。

**语法**

**express.Add(Description, Extensions, Position)**

*express*   一个代表 **FileDialogFilters** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Description* | 必选          | **String**   | 该文本表示要添加到筛选器列表中的文件扩展名的说明。           |
| *Extensions*  | 必选          | **String**   | 该文本代表要添加到筛选器列表中的文件扩展名。可以指定多个扩展名，每个扩展名必须以分号分隔。例如，可以向以下字符串分配参数：“*.txt; *.htm”。 注释 : 不需要在扩展名两侧添加括号。在说明和扩展名字符串连接成一个文件筛选器项时，WPS Office将自动在扩展名字符串两侧添加括号。 |
| *Position*    | 必选          | **Variant**  | 表示新控件在筛选列表中位置的数字。新筛选将插入到该位置的筛选之前。如果忽略该参数，筛选将添加到指定列表的末端。 |

**说明**

列表中的每个筛选器由两部分组成：文件扩展名（例如 .txt）和文件扩展名的文本说明（例如“文本文件”）。二者相结合，文件筛选器将在**“文件类型”**下拉列表框中显示为：文本文件(*.txt)。请注意，向列表中添加筛选器时，不会删除默认的筛选器。仅当选中**“窗口”**选项时才显示筛选器。如果 *Position* 无效，将显示超出范围运行时错误。如果 *Description* 和 *Extensions* 值无效，将显示运行时错误（分析）。

#### **FileDialogFilters.Clear**

删除当前在文件对话框中应用的所有筛选器。

**语法**

**express.Clear()**

*express*   一个代表 **FileDialogFilters** 对象的变量。

**示例**

下面的示例循环访问**“SaveAs”**对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for(let i = 1;i <= fdfs.Count;i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files         if(fdf.Extensions.indexOf("xls") > -1) {             MsgBox("Description of filter: " + fdf.Description)         }     }  }` |

#### **FileDialogFilters.Delete**

删除文件对话框筛选。

**语法**

**express.Delete(filter)**

*express*   一个代表 **FileDialogFilters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**         |
| -------- | ------------- | ------------ | ---------------- |
| *filter* | 可选          | **Variant**  | 要删除的筛选器。 |

#### **FileDialogFilters.Item**

获取属于指定的 **FileDialogFilters** 集合的成员的 **FileDialogFilter** 对象。

**语法**

**express.Item()**

*express*   一个代表 **FileDialogFilters** 对象的变量。

**成员属性**

#### **FileDialogFilters.Application**

获取一个 **Application** 对象，代表 **FileDialogFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **FileDialogFilters** 对象的变量。

#### **FileDialogFilters.Count**

获取一个 **Long** 类型的值，指示 **FileDialogFilters** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **FileDialogFilters** 对象的变量。

#### **FileDialogFilters.Creator**

获取一个 32 位整数，指示创建 **FileDialogFilters** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **FileDialogFilters** 对象的变量。

#### **FileDialogFilters.Parent**

获取 **FileDialogFilters** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FileDialogFilters** 对象的变量。

适用环境：web

适用平台：windows/linux