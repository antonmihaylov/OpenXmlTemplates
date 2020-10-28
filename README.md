
[![GitHub contributors](https://img.shields.io/github/contributors/Naereen/StrapDown.js.svg)](https://github.com/antonmihaylov/OpenXmlTemplates/graphs/contributors)
[![GitHub issues](https://img.shields.io/github/issues/Naereen/StrapDown.js.svg)](https://GitHub.com/antonmihaylov/OpenXmlTemplates/issues/)
[![License: LGPL v3](https://img.shields.io/badge/License-LGPL%20v3-blue.svg)](https://www.gnu.org/licenses/lgpl-3.0)


<!-- PROJECT LOGO -->
<br />
<p align="center">
  <!-- <a href="https://github.com/github_username/repo">
    <img src="images/logo.png" alt="Logo" width="80" height="80"/>
  </a>  -->

  <h2 align="center">Open XML Templates</h2>
  
  <p align="center">
    A .NET Standard Word documents templating system that doesn't need Word installed and is both designer and developer friendly
    <br />
    <br />
    <a href="https://github.com/antonmihaylov/OpenXmlTemplates/issues">Report Bug</a>
    ·
    <a href="https://github.com/antonmihaylov/OpenXmlTemplates/issues">Request Feature</a>
  </p>
</p>



<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installation](#installation)
* [Usage](#usage)
* [Supported tags](#supported-tags)
* [Roadmap](#roadmap)
* [Contributing](#contributing)
* [License](#license)
* [Contact](#contact)



<!-- ABOUT THE PROJECT -->
## About The Project

<div align="center">
    <img src="/ReadmeImages/screenshot_varaible_before.png?raw=true" width="60%"</img> 
</div>

With the library you can easily:
* Create word templates using only content controls and their tags
* Replace the content controls in a template with actual data from any source (json and a dictionary are natively supported)
* Repeat text based on a list (with nested variables and lists)
* Conditionaly remove text section
* Specify a singular and a plural word that should be used conditionaly, based of the length of a list

It is server-friendly, because it doesn't require Word installed. Only the Open XML Sdk is used for manipulating the document.

It is friendly to the designer of the document templates, because they don't need to have any coding skills and they won't have to write any
script-like snippets in the word document. Everything is instead managed by native Word content controls. 

### Built With

* [.NET Standard](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK)



<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running use one of the following methods:

 Install via nuget:

```
nuget install OpenXMLTemplates
```

 or clone the repo and reference OpenXMLTemplates.csproj in your project
 
```
git clone https://github.com/antonmihaylov/OpenXmlTemplates.git
```

<!-- USAGE EXAMPLES -->
## Usage

#### To create a template:

1. Open your document in Word
2. Open the *Developer* tab in the ribbon 
  (if you don't have it - open *File* tab, go to *Options* > *Customize Ribbon*.
  Under *Customize the Ribbon* and under *Main Tabs*, select the *Developer* check box.)
3. Under the *Controls* tab - add a new *Content Control* of your liking (*Plain text* is the simplest one - just text with formatting)
4. Select the newly added *Content control* and click *Properties* in the *Developer* ribbon
5. Change the *Tag* in the popup window to match one of the [supported tags](#supported-tags) (the tag name is case-insensitive - variable is the same as VARIABLE)

#### To create a document from a template, using the default content control replacers:

1. Create a new TemplateDocument. This represents your document and it neatly handles all
content controls in it, as well as the open/save/close file logic. Don't forget to call Dispose() on it
after you're done, or just use an "using" statement:
    ```c#
         using var doc = new TemplateDocument("path/to/my/document.docx");
    ```
2. Create a new VariableSource (currently available sources are a json string and a dictionary. 
You can also create your own class that implements IVariableSource). The variable source handles 
your data and extracts it in a way that the template engine can read it.
    ```c#
        var src = new VariableSource(jsonString); 
    ```
3. Create an OpenXmlTemplateEngine. A default one is provided (DefaultOpenXmlTemplateEngine). 
The default one contains all control replacers listed in the readme. You can disable/enable a control replacer by 
modifying the IsEnabled variable in it. You can also register your own replacer by callin RegisterReplacer on the engine.
    ```c#
        var engine = new DefaultOpenXmlTemplateEngine();
    ```
4. Call the ReplaceAll method on the engine using the document and the variable source
    ```c#
        engine.ReplaceAll(doc, src);
    ```
5. Save the edited document
    ```c#
       doc.SaveAs("result.docx"); 
   ```
 
 If you want to remove the content controls from the final document, but keep the content you have two options:
 1. Use the RemoveControlsAndKeepContent method on the TemplateDocument object 
 or 
 2. Set the KeepContentControlAfterReplacement boolean of the OpenXmlTemplateEngine

## Supported Tags

Note that if your variable names contain an underscore results may be unpredictable!
Note: to insert a new line, add a new line character (\r\n, \n\r, \n) in the data you provide, it will be parsed as a line break

### Variable

* Tag name: "variable_\<NAME OF YOUR VARIABLE\>" (the *variable* keyword is case-insensitive)
* Replaces the text inside the control with the value of the variable with the provided name
* Supports nested variable names (e.g. address.street)
* Supports array access (e.g. names[0])
* Supports nested variables using rich text content controls. For example: a rich text content control with
tag name address, followed by an inner content control with tag name variable_street is the same as variable.street
* Note that if you reference a variable from a nested control, that is available in the outer scope, but not in the inner scope - the outer scope variable will be used. 
* Supports variables inside repeating items, the variable name is relative to the repeated item.

  Example:
 
 - See example files in the [OpenXmlTemplatesTest/ControlReplacerTests/VariableControlReplacerTests folder](/OpenXMLTemplatesTest/ControlReplacersTests/VariableControlReplacerTests) and in
    the [OpenXmltemplatesTest/EngineTest folder](/OpenXMLTemplatesTest/EngineTest)
  
 ![](https://github.com/antonmihaylov/OpenXmlTemplates/blob/master/ReadmeImages/example_variable.png)
 
 
 ### Repeating

* Tag name: "repeating_\<NAME OF YOUR VARIABLE\>"  (the *repeating* keyword is case-insensitive)
* Repeats the content control as many times as there are items in the variable identified by the provided variable name.
* Complex fields with inner content controls are supported. Use the inner controls as you would normally, except
that the variable names will be relative to the list item. All default content controls can be nested.
* Note that if you reference a variable from a nested control, that is available in the outer scope, but not in the inner scope (the list item) - the outer scope variable will be used. That is useful if you want to include something in your list item's text output that is available in the global scope only.
* Add an inner content control with tag variable_index to insert the index of the current item (1-based)
* You can add extra arguments to the tag name (e.g. "repeating_\<VARIABLE NAME\>_extraparam1_extraparam2..."):
  * "separator_\<INSERT SEPARATOR STRING\>"- inserts a separator after each item (e.g. "repeating_\<VARIABLE NAME\>_separator_, " - this inserts a comma between each item)
  * "lastSeparator_\<INSERT SEPARATOR STRING\>"- inserts a special sepeartor before the last item (e.g. "repeating_\<VARIABLE NAME\>_separator_, _lastSeparator_and " - this inserts a comma between each item and an "and" before the last item)

  Example:
  
 
 - See example files in the [OpenXmlTemplatesTest/ControlReplacerTests/RepeatingControlTests folder](/OpenXMLTemplatesTest/ControlReplacersTests/RepeatingControlTests) and in
    the [OpenXmltemplatesTest/EngineTest folder](/OpenXMLTemplatesTest/EngineTest)
    
 ![](https://github.com/antonmihaylov/OpenXmlTemplates/blob/master/ReadmeImages/example_repeating.png)


### Conditional remove

* Tag name: "conditionalRemove_\<ENTER THE NAME OF YOUR VARIABLE\>"  (the *conditionalRemove* keyword is case-insensitive)
* Removes content controls based on the value of the provided variable
* If the variable value is evaluated to true (True, "true", 1, "1", non-empty list, non-empty dict) the control stays. If it doesn't - it is removed
* You can add extra arguments to the tag name (e.g. "conditionalRemove_\<VARIABLE NAME\>_extraparam1_extraparam2..."):
  * "OR" - applies an OR operation to the values. The control is removed if none of the values between the operator are true. (e.g. "conditionalRemove_\<VARIABLE NAME 1\>_or_\<VARIABLE NAME 2\>")
  * "EQ", "GT" and "LT" - checks if the value of the first variable equals ("eq"), is greather than ("gt") or is less than ("lt") the second variable's value. (e.g. "conditionalRemove_\<VARIABLE NAME 1\>_lt_\<VARIABLE NAME 2\>"). You can also provide a value to the operation, instead of a variable name (e.g. "conditionalRemove_\<VARIABLE NAME\>_lt_2). The control is removed if the supplied condition evaluates to false.
  * "NOT" - reverses the last value. (e.g. "conditionalRemove_\<VARIABLE NAME\>_not)
* You can also chain multiple arguments, e.g.  "conditionalRemove_\<VARIABLE NAME 1\>_not_or__\<VARIABLE NAME 2\>_and_\<VARIABLE NAME 3\>". Note that the expression is evaluated from left to right, with no recognition for the order of operations.

  Example:
  
   
 - See example files in the [OpenXmlTemplatesTest/ControlReplacerTests/ConditionalControlReplacerTest folder](/OpenXMLTemplatesTest/ControlReplacersTests/ConditionalControlReplacerTest) and in
    the [OpenXmltemplatesTest/EngineTest folder](/OpenXMLTemplatesTest/EngineTest)
    
  
 ![](https://github.com/antonmihaylov/OpenXmlTemplates/blob/master/ReadmeImages/example_conditionalRemove.png)


### Singular dropdown

* Works only with Dropdown content control!
* Tag name: "singular_\<ENTER THE NAME OF YOUR LIST VARIABLE\>"  (the *singular* keyword is case-insensitive)
* Replaces the text inside a content control with the appropriate value based on the length of the list variable with the provided name
* If the list variable has a length of 1 (or 0) the first value from the dropdown is used. If it's more than one - the second value from the dropdown is used.

  Example:
 
    
 - See example files in the [OpenXmlTemplatesTest/ControlReplacerTests/DropdownControlReplacersTests/SingularsTest folder](/OpenXMLTemplatesTest/ControlReplacersTests/DropdownControlReplacersTests/SingularsTest)
    
  
 ![](https://github.com/antonmihaylov/OpenXmlTemplates/blob/master/ReadmeImages/example_singular.png)
 
 
### Conditional dropdown

* Works only with Dropdown content control!
* Tag name: "conditional_\<ENTER THE NAME OF YOUR LIST VARIABLE\>"  (the *conditional* keyword is case-insensitive)
* Replaces the text inside a content control with the appropriate value based on the length of the variable with the provided name
* If it's evaluated to true (aka is true, "true", 1, "1", non-empty list, non-empty dict) - the first value from the dropdown is used. If it's not - the second value is used.
* You can use the same extra arguments as in the Conditional remove replacer 

 - See example files in the [OpenXmlTemplatesTest/ControlReplacerTests/DropdownControlReplacersTests/ConditionalDropdownControlReplacerTest folder](/OpenXMLTemplatesTest/ControlReplacersTests/DropdownControlReplacersTests/ConditionalDropdownControlReplacerTest)


<!-- ROADMAP -->
## Roadmap

See the [open issues](https://github.com/github_username/repo/issues) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



<!-- LICENSE -->
## License

Distributed under the LGPLv3 License. See `LICENSE` for more information.



<!-- CONTACT -->
## Contact

Anton Mihaylov - antonmmihaylov@gmail.com

Project Link: [https://github.com/antonmihaylov/OpenXmlTemplates](https://github.com/antonmihaylov/OpenXmlTemplates)



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/othneildrew/Best-README-Template.svg?style=flat-square
[contributors-url]: https://github.com/othneildrew/Best-README-Template/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/othneildrew/Best-README-Template.svg?style=flat-square
[forks-url]: https://github.com/othneildrew/Best-README-Template/network/members
[stars-shield]: https://img.shields.io/github/stars/othneildrew/Best-README-Template.svg?style=flat-square
[stars-url]: https://github.com/othneildrew/Best-README-Template/stargazers
[issues-shield]: https://img.shields.io/github/issues/othneildrew/Best-README-Template.svg?style=flat-square
[issues-url]: https://github.com/othneildrew/Best-README-Template/issues
[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=flat-square
[license-url]: https://github.com/othneildrew/Best-README-Template/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/othneildrew
