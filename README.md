# hide_indent
LibreOffice Calc Python macro for tree node expand and collapse. (By treating indentations by empty cells/space characters/other characters as indication for parent/child relationship for rows.)

Inspired by emacs org-mode's outline folding and unfolding. Similar to my other project on github named group_indent but this one don't have the limitation of Office's maximum nested group level of seven. 

## How to use
Select the cell on the row you want to expand/collapse (the column of the cell has to be located on the same column as the root node) and then call the hide_selection() function.

## Screenshots
1. Select the row for it sub branch to collapse and then call hide_selection(). 
![Alt text](images/expand_all.png)

2. All sub nodes are collpased now. Call hide_selection() one more time.
![Alt text](images/collapse_all.png)

3. Only the 1st level of sub nodes are expanded. Call hide_selection() one more time will expand all nodes.
![Alt text](images/expand_one_lvl.png)

