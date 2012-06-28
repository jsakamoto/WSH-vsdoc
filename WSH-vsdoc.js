var WScript = {
    Echo: function (text) { },
    Quit: function () { },
    ScriptFullName: ""
};

var ActiveXObject = function (clsid) {
    ///<param name="clsid">
    // Specify CLSID of COM object to create. For example:
    ///<para>"Scripting.FileSystemObject"</para>
    ///<para>"WScript.Shell"</para>
    ///</param>
    if (clsid == "Scripting.FileSystemObject") {

        var File = function () {
            this.Name = "";
            this.Size = 0;
        };

        var Folder = function () {
            this.Name = "";
            this.Size = 0;
            this.Files = [new File()];
            this.Move = function (destinationPath) { };
        };
        Folder.prototype.SubFolders = [new Folder()];

        this.GetFolder = function (fullPath) { return new Folder(); };
        this.FolderExists = function (fullPath) { return true; };
        this.MoveFolder = function (sourceFullPath, destinationFullPath) { };
        this.CreateFolder = function (fullPath) { };
    }

    if (clsid == "WScript.Shell") {
        this.Run = function (command, windowStyle, waitOnReturn) {
            ///<param name="command" type="String"></param>
            ///<param name="windowStyle" type="int">optional. 0:hidden, 1:restore, 2:minimize, 3:maximize, ...</param>
            ///<param name="waitOnReturn" type="bool">optional.</param>
        }

        // for Registry
        // http://msdn.microsoft.com/en-us/library/x05fawxd.aspx
        this.RegRead = function (strName) {
            /// <param name="strName" type="String">String value indicating the key or value-name whose value you want.</param>
        }
        // http://msdn.microsoft.com/en-us/library/yfdfhz1b.aspx
        this.RegWrite = function (strName, anyValue, strType) {
            /// <param name="strName" type="String">String value indicating the key-name, value-name, or value you want to create, add, or change.</param>
            /// <param name="anyValue" type="Object">The name of the new key you want to create, the name of the value you want to add to an existing key, or the new value you want to assign to an existing value-name.</param>
            /// <param name="strType" type="String">Optional. String value indicating the value's data type, REG_SZ, REG_DWORD, REG_BINARY, REG_EXPAND_SZ. The REG_MULTI_SZ type is not supported for the RegWrite method.</param>
        }
    }
};

var Enumerator = function (enumerable) {
    this.atEnd = function () { return false; };
    this.moveNext = function () { };
    this.item = function () { return enumerable[0]; };
};
