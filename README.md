# Visio Parsing Tool

## Introduction
This is the POC mock for a tool designed to generate and execute test cases for QA purposes. This version has some neat graph algorithms showing how different elements of a Visio file can be parsed to collect interesting data and how that data can be analyzed.

The goal of this version of the tool is to aid in the design and usage of Visio files for streamlining QA operations, namely test case handling, by providing a large amount of data collection and some optimal cases to use for streamlining IVR testing. The POC here will be expanded into a full application in a later version, and that version will feed its collected data into TestRail to fully generate and execute a large number of test cases

## Usages
The current version of the Visio Parser can:
- Prompt the user to ask which Visio document to select if there are multiple or autoselect if there is only one provided
- Unzip and parse extracted XML contents from Visio files
- Calculate the starting and ending nodes in each page, providing a list of potential options before user configuration
- Traverse every shape, connection, on/off-page references (with specification) to build a multi-page graph
- Generate the permutations (paths from start to end) and display them for the user
- Determine the optimal minimum paths needed to cover all edges / cases to simplify QA testing and display them for the user
- Allow the user various ways to specify how start and end nodes should be selected (Text, shape, indiscriminately, etc.)
- Allow the user to specify how on-page and off-page references can be identified
- Provide the option to rezip the exracted contents with IDs overwritten on text for ID mapping on the permutations
- Determines the optimal minimum paths needed to cover all edges / cases to simplify QA testing
- Handle infinite loops from cyclical references to ensure progress during runtime

## TDL
### Future Implementation
- Enable function return detection, i.e traversing to an off-page reference, then popping back to the previous page when encountering a node that has some kind of identifier for "return to calling function"
- Add generate a unique id for each shape, the ids for each shape are unique to a shape only among the other shapes on the page, but when dealing with flows across multiple pages there may be identification issues
- Develop test cases using paths on a page-by-page basis, should help modularize tests in a more performance-effective way as far less runtime is wasted among large off-page reference cycles
- Investigate a useful format and general methods of development for test case creation to facilitate ingestion into TestRail/Cyara
- Develop a UI for easier config options and possibly reduce the number of config options needed to enter
	- possibly combine the text/id option into one and determine what the user is entering based on the read-in C# datatype being convertable to an int
 	- could possibly automatically join on-page references rather than asking for an id

### Done:
- Tracked number of paths per page to calculate number of test cases to write
- Added functionality to remove any vertex with no incoming or outgoing edges to reduce clutter
- Converted all the strings representing shapes (using the ID) into XML objects
- Cleaned up and functionalized the program
- Mapping (change all text in the xml to the ID and recompress it)
- Printed only the necessary paths in text to save on performance and storage
- Cleaned up what goes into the output file to be useful for searching data
- Option menu implemented to enable user to specify runtime parameters and add configurations
- Functionality to check start nodes for text or a specific master shape (can use config menu)
- Functionality to check if a starting node has a path that is contained within another path (expensive computationally but implemented as minimum paths)
- Implemented the usage of off-page references in the path determination
  - design changes necessary to implement, will need to make one large graph with added edges between off-page references
  - the off-page references will be specified by the user and parsed separately after the normal page graphs via hyperlinks for page references and text for checkpoints
- Implemented the usage of on-page references in the path determination
  - using much of the same logic as the off-page references but by matching the text of shapes with provided master IDs on the same page and determining which shape points to which
- Improved permutation/minimum path algorithms for optimization (20-minute runtime on test case down to 20-second runtime)
- Changed filepaths to in the callflow handler work universally rather than hardcoding dev file path
- Implemented a check for off-page references in the xml data to automatically connect them (if no receiving off-page references on the page, then create an edge to the start node of that page instead)
- Implemented cycle detection to prevent infinite loops
- Reduced memory constraints to avoid stack overflow during recursion by making DFS iterative instead of recursively filling up stack memory
- Moved parsing to be done upfront so that specifying node information doesn't require two separate program executions
- Implemented node proportion calculation to determine how much of the graph was parsed / is unreachable

  
## Installation
1. Download Visual Studio if it is not currently installed
2. Clone the repository to download the contents to your local machine
3. Open up the downloaded repo and double click on the .sln file to open up the solution in Visual Studio

## Getting Started
Currently, usage of the parser is somewhat convoluted without a dedicated interface for the various different configurations. In order to parse the Visio accurately, the parser needs some additional information provided by the user, or else it wouldn't be able to tell which nodes are start nodes and which nodes are references. To assist with that, the user can enter in some information such as different IDs for various shapes in the Visio. To get started, follow these steps:
1. Follow the instructions in the Installation portion
2. Navigate to the Documents folder within the repository directory (VisioParse.ConsoleHost\VisioParse.ConsoleHost\Documents), copy the desired Visio into this file
3. Open up the solution by clicking on the .sln file
4. Navigate to the Solution Explorer tab on the left, click the dropdown on the Documents folder to see the Visio inputted previously
5. Right click on the Visio, click Properties, then select Copy if newer
6. Run the program
7. If only one Visio document was placed into the documents folder, the program will automatically select that file for parsing, else if there are more it will ask the user to select the desired one
8. The internal specifications of various important shapes will be provided in the console, i.e. contents of a "Legend" page or identified potential start/end nodes and on-page references
9. To parse with specification, follow the menu dialog to guide the parser to pick true identifiers (either text or ids) for start and end nodes
10. Then, assist the parser with any potential references identification, off-page references are handled automatically as long as the hyperlink is functional
11. Allow the parser to run, it will provide outputs for every combination of start and end nodes it attempts to find paths between
12. After the paths are calculated, it will crunch them down into the necessary minimum paths needed to cover every vertex, i.e. ideal test cases
13. The output files can be found at VisioParse.ConsoleHost\VisioParse.ConsoleHost\bin\Debug\net7.0, pageInfo_filename contains info for each shape by page, paths contains the exhaustive list of every possible path while limited cycles, minPaths contains the shortened list of ideal paths for test cases
14. For an easier inspection of path information, choose to rezip the files back into a Visio, then open up the Visio from the same bin directory as the other output files and check the Master IDs

There are additional instructions for installation/execution/usage in the documentation included

## Documentation
There is additional documentation that is updated after major changes to the program which is included in this repository
