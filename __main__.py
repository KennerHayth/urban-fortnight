import Contractor_Sorting
import Remove_Assignments
import pandas as pd

pending_changes = Contractor_Sorting.main()

Remove_Assignments.main(pending_changes)