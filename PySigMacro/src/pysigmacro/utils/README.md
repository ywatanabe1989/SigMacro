<!-- ---
!-- Timestamp: 2025-03-25 05:36:56
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/README.md
!-- --- -->

## Renaming Example

``` python
import numpy as np
import pandas as pd

import pysigmacro as ps

# PARAMS
PLOT_TYPE = "line"
CLOSE_OTHERS = True
PATH = ps.path.copy_template("line", rf"C:\Users\wyusu\Downloads")
DF = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

spw = ps.con.open(PATH)
notebooks = spw.Notebooks_obj
# print(notebooks.list)
notebook = notebooks[notebooks.find_indices(f"{PLOT_TYPE}")[0]]

# From here, templates defines indices and names
notebookitems = notebook.NotebookItems_obj
graphitem_s = notebook.NotebookItems_obj[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
]
graphitem_m = notebook.NotebookItems_obj[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_M")[0]
]
graphitem_l = notebook.NotebookItems_obj[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_L")[0]
]

ps.utils.run_macro(
    graphitem_s, "RenameXYLabels_macro", xlabel="X Label 1", ylabel="Y Label 1"
)
ps.utils.run_macro(
    graphitem_m, "RenameXYLabels_macro", xlabel="X Label", ylabel="Y Label"
)
ps.utils.run_macro(
    graphitem_l, "RenameXYLabels_macro", xlabel="X Label 2", ylabel="Y Label 2"
)
```


# EOF

<!-- EOF -->