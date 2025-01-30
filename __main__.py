import Contractor_Sorting
import Remove_Assignments
import pandas as pd

pending_changes = Contractor_Sorting.main()

pending_changes = pd.DataFrame(pending_changes, columns=["Email"])

Remove_Assignments.main(pending_changes)