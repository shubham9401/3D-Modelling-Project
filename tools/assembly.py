
import pandas as pd
marks = [1, 25, 50, 75, 100]
marks_obtained = [90, 58, 32, 87, 24]
categories = ['Poor', 'Average', 'Good', "Excellent"]
stats = pd.cut(marks_obtained, marks, labels=categories)

print(pd.value_counts(stats))