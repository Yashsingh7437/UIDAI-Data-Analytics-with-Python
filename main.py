import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

data = pd.read_excel("data.xlsx", skiprows=2)
data = data.iloc[:,1:]
data =  data.set_index("Date")

# print(data)


data['Supplement'] = data['Supplement'].map({"Yes":1,"No":0})
data['Journaling'] = data['Journaling'].map({"Yes":1,"No":0})
# plt.plot(data.index,data['Workout'])
# plt.title("Workout Stats")
# plt.xlabel("Date")
# plt.ylabel("Workout Time (in Minutes)")
# plt.show()

print(data.corr())
sns.heatmap(data.corr(),annot= True)
plt.show()