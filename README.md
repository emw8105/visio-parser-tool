# Visio Parsing Tool

## Introduction
This is the POC mock for a tool designed to generate and execute test cases for QA purposes. This version simply has some neat graph algorithms that show off how different elements of a Visio file can be parsed to collect interesting data.

The goal of this version of the tool is to aid in the design and usage of Visio files for streamlining QA operations, namely test case handling, by providing a large amount of data collection. 

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
### Done:
- track number of paths per page to calculate number of test cases to write
- don't add any vertex with no incoming or outgoing edges to reduce clutter
- convert all the strings representing shapes (using the ID) into XML objects
- clean up and functionalize the program
- mapping (change all text in the xml to the ID and recompress it)
- also print out the paths as text for quick reference
- clean up what goes into the output file to be useful for searching data
- option menu implemented to enable user to specify runtime parameters and add configurations
- check start nodes for text or a specific master shape (can use option menu)
- can check if a starting node has a path that is contained within another path (expensive computationally but implemented as minimum paths)
- implement the usage of off-page references in the path determination
  - design changes necessary to implement, will need to make one large graph with added edges between off-page references
  - the off-page references will be specified by the user and parsed separately after the normal page graphs via hyperlinks for page references and text for checkpoints
- implement the usage of on-page references in the path determination
  - using much of the same logic as the off-page references but by matching the text of shapes with provided master IDs on the same page and determining which shape points to which
### WIP:
- create guidelines or a template for example usage that the tool can handle, likely whatever excel supports - WIP, updating every so often with new info
### Stretch Goals:
- implement return logic
  - i.e. an off-page reference traverses to another flow which then returns to the calling location in the flow
- figure out how to parse .yaml files for genesys architect and see if it can be recreated in visio? - not ideal to pursue because it would generate faulty test cases from faulty arcitect flows

## Installation
1. Download Visual Studio if it is not currently installed
2. Clone the repository to download the contents to your local machine
3. Open up the downloaded repo and double click on the .sln file to open up the solution in Visual Studio

There are additional instructions for installation/execution/usage in the documentation included

## Documentation
There is additional documentation that is updated after major changes to the program which is included in this repository
