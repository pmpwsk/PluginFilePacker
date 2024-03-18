# PluginFilePacker
Visual Studio extension (VSIX) that generates a [WebFramework](https://uwap.org/projects/webframework) plugin file handler (FileHandler.cs) for the files in Files/.

Example of a generated file on GitHub: [ServerPlugin/FileHandler.cs](https://github.com/pmpwsk/ServerPlugin/blob/master/FileHandler.cs)

Website: https://uwap.org/projects/plugin-file-packer

Version for Visual Studio Code: https://marketplace.visualstudio.com/items?itemName=uwap-org.uwap-pluginfilepacker-vsc

## Main features
- Generating a plugin file handler for static files (using text for dynamically modified text files or using a .RESX file / base64 string otherwise (depending on the setting))
- Calling a custom file handler if no matching file was found (see "Usage" below)
- Automatically detecting the namespace of the existing FileHandler.cs or FileHandlerCustom.cs
- Dynamically replacing [PATH_PREFIX] with the current plugin path prefix, [PATH_HOME] with the plugin's current home/root path, [DOMAIN] with the current domain and [DOMAIN_MAIN] with the main domain for the current domain

## Installation
You can find this extension on the Visual Studio marketplace as: [PluginFilePacker](https://marketplace.visualstudio.com/items?itemName=uwap-org.uwap-PluginFilePacker)

Alternatively, you can download the .vsix file from the [releases on GitHub](https://github.com/pmpwsk/PluginFilePacker/releases) and open it to install the extension.

## Usage
To execute the command, go to: <code>Tools > Generate FileHandler.cs</code> (if there are multiple projects in your solution, you might need to click on the desired project to select it!)

There are some settings under <code>Tools > Options > PluginFilePacker</code>, those are documented in their description.

The generated partial class will have the name of the project, so make sure the rest of your plugin matches that.

To create a custom file handler as a fallback, create a file <code>FileHandlerCustom.cs</code> using the same namespace and partial class, then add two methods like this:<br/><code>byte[]? GetFileCustom(string relPath, string pathPrefix, string domain)</code><br/><code>string? GetFileVersionCustom(string relPath)</code>
