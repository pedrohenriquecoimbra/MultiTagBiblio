# MultiTagBiblio

This program is thought as a tool to analyse a set of scientific papers read with Zotero under different scopes.

After text highlighting and commenting, you end up with information blocs (referenced sentences) and associated notes taken during reading. The extraction of these elements is automatic, thus you will be asked to provide this program your Zotero installation directory and the targetted collection. This program enables manual multitagging according to a user-built plan. Notes can be inserted in each plan category to summarize tagged information blocs.

An experimental semantic classification of all information blocs is also proposed for new topics classification to emerge from a given bibliography.

## Installation

- install git if not already done : https://git-scm.com/downloads

- install python if not already done : https://docs.conda.io/en/latest/miniconda.html

- open terminal (search **cmd** for Windows users). Go to your target directory with : *cd C:\path\to\your\directory*

- execute : *git clone https://github.com/MTreestanG/MultiTagBiblio*

## Usage

Double click *MultiTagBiblio.bat* for Windows users (or just execute *MultiTagBiblio.py* from terminal)

This program is meant to use your Zotero local database as a starting point. On first use, point to your Zotero installation directory (usually C:\Users\yourname\Zotero) and mention the name of the parent collection in Zotero in which your papers are grouped. A proposition to create a shortcut will also be made. 

On launch, click the *Import Zotero* button to get your last annotations. Please note Zotero should be closed to perform this action.

Then build a plan(s) related to your research questions in the left pane :
Click the *Add* button to add a plan category. You can select a position in your existing plan to position the new element easily. Other button commands are comprehensive and enable structure modifications.

You can then tag whole article by selecting it and press *Tagging*, or just one block when modifying afterwards, with the same button. While tagging, follow the *Shell* pane title instructions and press *Next* once you have selected plan categories related to the current displayed sentence.

Finally, when selecting a plan category, press *Take notes* button to start taking notes in the bottom text entry of the window. Don't forget to press *Save* before switching to another.

## Visualization

You can use the *Search* button field to search all sentences from all articles. Search term will be fetched its lemma to extend search results.

(Experimental) You can use the *Topics* button to classify all sentences according to Hierarchical Ascending Classification (vectorized with *sentence_transformers* package). Witness the adapted number of categories from the plot and enter the corresponding value to get groups displayed and the most occuring topics in these groups.

## Update

This is work in progress. Feel free to report a bug or request features.

If connected to network, your program will update on start.
