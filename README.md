# Visio Parsing Tool

## Introduction
This is the POC mock for a tool designed to generate and execute test cases for QA purposes. This version has some neat graph algorithms showing how different elements of a Visio file can be parsed to collect interesting data and how that data can be analyzed.

The goal of this version of the tool is to aid in the design and usage of Visio files for streamlining QA operations, namely test case handling, by providing a large amount of data collection and some optimal cases to use for streamlining IVR testing. The POC here will be expanded into a full application in a later version, and that version will feed its collected data into TestRail to fully generate and execute a large number of test cases

## Usages
The current version of the Visio Parser can:
- Unzip and parse extracted XML contents from Visio files
- Calculate the starting and ending nodes in each page
- Traverse every shape, connection, on/off-page references (with specification) to build a multi-page graph
- Generate the permutations (paths from start to end) and display them for the user
- Determine the optimal minimum paths needed to cover all edges / cases to simplify QA testing and display them for the user
- Allow the user various ways to specify how start and end nodes should be selected (Text, shape, indiscriminately, etc.)
- Allow the user to specify how on-page and off-page references can be identified
- Provide the option to rezip the exracted contents with IDs overwritten on text for ID mapping on the permutations
- Determines the optimal minimum paths needed to cover all edges / cases to simplify QA testing

## TDL
### WIP:
- prevent infinite loops from cyclical off-page references, find some way to direct permutations to avoid the same sequence of operations from happening in a cycle, could take advantage of page property to avoid checking every previously created path
  - could also possibly check if the current node has already been included if it is an off-page reference, if so then backtrack instead of taking the reference, investigate solutions that dont jeapordize possible viable paths
- enable function return detection, i.e traversing to an off-page reference, then popping back to the previous page when encountering a node that has some kind of identifier for "return to calling function"
- streamline the process for retrieving IDs, parse the Visio information and make it available before determining node ids and stuff
	- also, possibly have it displayed in the console, i.e. look for a page titled "Legend" and print the ids of all the components on that page
	- add a console detection for potentially missed start and end nodes, i.e. for the query that searches for master id, check if any not in that category have text that matches "Start" or "End" or "Disconnect", etc. and print those
- investigate possible improved loading requirements, maybe determine how many documents are in a folder, if it's only one then do that, else ask for the title idk
- use a better design such as this pseudocode to allow file prints to be added while parsing so it's not an all-or-nothing solution (better debugging)

public static IEnumerable<Vertex> YieldingMethod()
{
  foreach(outer in outers)
    foreach(inner in inners)
    {
      foreach(vertex in ThingThatReturnsVertices())
        yield return vertex;
    }
}

foreach(var vertex in YieldingMethod())
  AppendToFile(vertex);

- memory requirements might need some scrutiny but judge this after infinite loops have been solved
- rewrite readme to include these components once finished (initialization) and incorporate usage documentation
### Done:
- Track number of paths per page to calculate number of test cases to write
- Don't add any vertex with no incoming or outgoing edges to reduce clutter
- Convert all the strings representing shapes (using the ID) into XML objects
- Clean up and functionalize the program
- Mapping (change all text in the xml to the ID and recompress it)
- Also print out the paths as text for quick reference
- Clean up what goes into the output file to be useful for searching data
- Option menu implemented to enable user to specify runtime parameters and add configurations
- Check start nodes for text or a specific master shape (can use option menu)
- Check if a starting node has a path that is contained within another path (expensive computationally but implemented as minimum paths)
- Implement the usage of off-page references in the path determination
  - design changes necessary to implement, will need to make one large graph with added edges between off-page references
  - the off-page references will be specified by the user and parsed separately after the normal page graphs via hyperlinks for page references and text for checkpoints
- Implement the usage of on-page references in the path determination
  - using much of the same logic as the off-page references but by matching the text of shapes with provided master IDs on the same page and determining which shape points to which
- Improve permutation/minimum path algorithms for optimization (20-minute runtime on test case down to 20-second runtime 😎)
- Change filepaths to in the callflow handler work universally rather than hardcoding dev file path
- To save visio development time, implement a check for off-page references (if no receiving off-page references on the page, then create an edge to the start node of that page instead)
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
6. Open up CallflowHandler.cs
7. Scroll down to the CallflowHandler constructor, in the line where the FileName is set to a string, change it to be the name of the Visio file (including the .vsdx)
8. Run the program
9. If knowledge of various IDs are not already known, parse indiscriminately
10. Output files can be found at VisioParse.ConsoleHost\VisioParse.ConsoleHost\bin\Debug\net7.0, check the contents of pageInfo_filename to see the Master IDs of various shapes
11. For an easier inspection, choose to rezip the files back into a Visio, then open up the Visio from the same bin directory as the other output files and check the Master IDs
12. Take note of the starting shape Master IDs, ending shape Master IDs, and on-page reference Master IDS
13. When ready, rerun the program and instead choose to parse with specification
14. Follow the menu prompts and enter corresponding IDs when prompted
15. Once the parser finishes execution, inspect the output files (especially the minPaths_filename file) to analyze the outputs

There are additional instructions for installation/execution/usage in the documentation included

## Documentation
There is additional documentation that is updated after major changes to the program which is included in this repository
