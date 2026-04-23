# ============================================================
# PROJECT 2: HR EMPLOYEE ATTRITION ANALYSIS
# Run this file in Jupyter Notebook or any Python environment
# Required: pip install pandas matplotlib seaborn openpyxl
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ── LOAD DATA ────────────────────────────────────────────────
df = pd.read_excel("P2_HR_Attrition.xlsx", sheet_name="Employee_Data", skiprows=1)
print("Data loaded:", df.shape)
print(df.head())

# ── BASIC CHECK ──────────────────────────────────────────────
print("\nAttrition Count:")
print(df['Attrition'].value_counts())
print("\nAttrition Rate:", round(df['Attrition'].eq('Yes').mean() * 100, 1), "%")

# ── CHART 1: Attrition Rate by Department ────────────────────
dept = (
    df.groupby('Department')['Attrition']
    .apply(lambda x: round((x == 'Yes').sum() / len(x) * 100, 1))
    .reset_index()
)
dept.columns = ['Department', 'Attrition_Rate']
dept = dept.sort_values('Attrition_Rate', ascending=False)

plt.figure(figsize=(9, 5))
bars = plt.bar(dept['Department'], dept['Attrition_Rate'],
               color=['#C00000' if v > 40 else '#ED7D31' if v > 25 else '#70AD47'
                      for v in dept['Attrition_Rate']])
plt.axhline(y=dept['Attrition_Rate'].mean(), color='navy',
            linestyle='--', linewidth=1.5, label='Company Average')
plt.title('Attrition Rate by Department (%)', fontsize=14, fontweight='bold', pad=15)
plt.ylabel('Attrition Rate (%)', fontsize=11)
plt.xlabel('Department', fontsize=11)
plt.legend()
for bar, val in zip(bars, dept['Attrition_Rate']):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
             f'{val}%', ha='center', fontsize=10, fontweight='bold')
plt.tight_layout()
plt.savefig('chart1_dept_attrition.png', dpi=150, bbox_inches='tight')
plt.show()
print("Chart 1 saved!")

# ── CHART 2: Salary — Left vs Stayed ─────────────────────────
plt.figure(figsize=(8, 5))
colors = {'No': '#2E75B6', 'Yes': '#C00000'}
for att, grp in df.groupby('Attrition'):
    plt.hist(grp['Monthly_Salary'], bins=8, alpha=0.6,
             label=f"{'Stayed' if att=='No' else 'Left'} ({len(grp)})",
             color=colors[att])
plt.title('Monthly Salary Distribution: Left vs Stayed', fontsize=14,
          fontweight='bold', pad=15)
plt.xlabel('Monthly Salary (PKR)', fontsize=11)
plt.ylabel('Number of Employees', fontsize=11)
plt.legend(fontsize=11)
plt.tight_layout()
plt.savefig('chart2_salary_attrition.png', dpi=150, bbox_inches='tight')
plt.show()
print("Chart 2 saved!")

# ── CHART 3: Attrition by Age Group ──────────────────────────
df['Age_Group'] = pd.cut(df['Age'], bins=[20, 28, 35, 45, 55],
                          labels=['20-28', '29-35', '36-45', '46-55'])
age_att = (
    df.groupby('Age_Group')['Attrition']
    .apply(lambda x: (x == 'Yes').sum())
    .reset_index()
)
age_att.columns = ['Age_Group', 'Left']

plt.figure(figsize=(8, 5))
plt.bar(age_att['Age_Group'].astype(str), age_att['Left'],
        color=['#C00000', '#ED7D31', '#FFC000', '#70AD47'])
plt.title('Employees Who Left — by Age Group', fontsize=14,
          fontweight='bold', pad=15)
plt.ylabel('Number of Employees Who Left', fontsize=11)
plt.xlabel('Age Group', fontsize=11)
for i, val in enumerate(age_att['Left']):
    plt.text(i, val + 0.1, str(val), ha='center', fontsize=12, fontweight='bold')
plt.tight_layout()
plt.savefig('chart3_age_attrition.png', dpi=150, bbox_inches='tight')
plt.show()
print("Chart 3 saved!")

# ── CHART 4: Overtime vs Attrition ───────────────────────────
ot = df.groupby(['Overtime', 'Attrition']).size().unstack(fill_value=0)
ot.plot(kind='bar', figsize=(7, 5), color=['#2E75B6', '#C00000'],
        edgecolor='white', width=0.5)
plt.title('Overtime vs Attrition', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Overtime', fontsize=11)
plt.ylabel('Number of Employees', fontsize=11)
plt.legend(['Stayed', 'Left'], fontsize=11)
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('chart4_overtime_attrition.png', dpi=150, bbox_inches='tight')
plt.show()
print("Chart 4 saved!")

# ── CHART 5: Job Satisfaction vs Attrition ───────────────────
sat = df.groupby(['Job_Satisfaction', 'Attrition']).size().unstack(fill_value=0)
sat.plot(kind='bar', figsize=(8, 5), color=['#2E75B6', '#C00000'],
         edgecolor='white', width=0.5)
plt.title('Job Satisfaction vs Attrition', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('Job Satisfaction Score (1=Low, 5=High)', fontsize=11)
plt.ylabel('Number of Employees', fontsize=11)
plt.legend(['Stayed', 'Left'], fontsize=11)
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('chart5_satisfaction_attrition.png', dpi=150, bbox_inches='tight')
plt.show()
print("Chart 5 saved!")

# ── SUMMARY STATISTICS ───────────────────────────────────────
print("\n" + "="*50)
print("KEY FINDINGS SUMMARY")
print("="*50)
left = df[df['Attrition'] == 'Yes']
stayed = df[df['Attrition'] == 'No']
print(f"Total Employees: {len(df)}")
print(f"Employees Left: {len(left)} ({round(len(left)/len(df)*100,1)}%)")
print(f"Avg Salary - Left:   PKR {int(left['Monthly_Salary'].mean()):,}")
print(f"Avg Salary - Stayed: PKR {int(stayed['Monthly_Salary'].mean()):,}")
print(f"Overtime & Left: {len(left[left['Overtime']=='Yes'])} out of {len(left)}")

dept_att = df.groupby('Department')['Attrition'].apply(
    lambda x: f"{(x=='Yes').sum()}/{len(x)} ({round((x=='Yes').mean()*100,1)}%)"
)
print("\nAttrition by Department:")
print(dept_att.to_string())

# ── SAVE RESULTS TO EXCEL ────────────────────────────────────
with pd.ExcelWriter('P2_Python_Results.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Full_Data', index=False)
    dept.to_excel(writer, sheet_name='Dept_Attrition', index=False)
    age_att.to_excel(writer, sheet_name='Age_Attrition', index=False)
    summary = pd.DataFrame({
        'Metric': ['Total Employees', 'Left', 'Stayed', 'Attrition Rate %',
                   'Avg Salary Left', 'Avg Salary Stayed'],
        'Value': [len(df), len(left), len(stayed),
                  round(len(left)/len(df)*100,1),
                  int(left['Monthly_Salary'].mean()),
                  int(stayed['Monthly_Salary'].mean())]
    })
    summary.to_excel(writer, sheet_name='Summary_Stats', index=False)

print("\n✅ All charts saved + Excel results exported to P2_Python_Results.xlsx")
