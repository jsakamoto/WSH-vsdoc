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
    }
};

var Enumerator = function (enumerable) {
    this.atEnd = function () { return false; };
    this.moveNext = function () { };
    this.item = function () { return enumerable[0]; };
};
