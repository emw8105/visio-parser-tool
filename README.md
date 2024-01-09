# Visio Parsing Tool

## Introduction
This is the POC mock for a tool designed to generate and execute test cases for QA purposes. This version simply has some neat graph algorithms that show off how different elements of a Visio file can be parsed to collect interesting data.

The goal of this version of the tool is to aid in the design and usage of Visio files by providing a large amount of data collection. 

## Usages
The current version of the Visio Parser can:
- Unzip and parse extracted XML contents from Visio files
- Calculate the starting and ending nodes in each page
- Generate the permutations (paths from start to end) and display them for the user
- Allow the user various ways to specify how start and end nodes should be selected (Text, shape, indiscriminately, etc.)
- Provide the option to rezip the exracted contents with IDs overwritten on text for ID mapping on the permutations
- Determines the optimal minimum paths needed to cover all edges / cases to simplify QA testing

## TDL
- track number of paths per page to calculate number of test cases to write - done
- don't add any vertex with no incoming or outgoing edges to reduce clutter - done
- convert all the strings representing shapes (using the ID) into XML objects - done
- clean up and functionalize the program - done
- mapping (change all text in the xml to the ID and recompress it?) - done
- also print out the paths as text for quick reference - done
- clean up what goes into the output file to be useful for searching data - outputs split to categorize them, done
- option menu implemented to enable user to specify runtime parameters and add configurations - done
- check start nodes for text or a specific master shape (can use option menu) - done
- implement the usage of off-page references in the path determination - WIP
  - would make the graph an array of graphs with indexes as page numbers, when a shape has a certain master,then could scan the text to see the location to jump to
  - would need to make enabling the references a configuration option where the master ID would need to be specified and the text follows some format
- can check if a starting node has a path that is contained within another path (could be expensive computationally)
- figure out how to parse .yaml files for genesys architect and see if it can be recreated in visio
- consider any other data that would be useful to parse and perform some algorithm on
- create guidelines or a template for example usage that the tool can handle, likely whatever excel supports - WIP

## Installation
1. Download Visual Studio if it is not currently installed
2. Clone the repository to download the contents to your local machine
3. Open up the downloaded repo and double click on the .sln file to open up the solution in Visual Studio

There are additional instructions for installation/execution/usage in the documentation included

## Documentation
There is additional documentation that is updated after major changes to the program which is included in this repository
