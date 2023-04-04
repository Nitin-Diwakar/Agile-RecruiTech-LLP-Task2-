#import essential library
import pandas as pd

# reading the given excel file
xls = pd.ExcelFile("Data Science Trainer Assignment.xlsx")

'''======================================================================
 Note: The below reading sheet i.e. input 1-2 this type because data is small
 but if the table present is large and other more information is also present
 at that sheet then for that suitation we have to copy the table and paste
it to the new sheet to process further.
===================================================================== '''

# exporting the given input-1 in the variable(i.e input1)
# before that while reading the file adding some extra paratmeter becuase different
# information present in single sheet and we only need the input table
input1 = pd.read_excel(xls,"(Input) User IDs",skiprows=10,usecols=["Name","Team Name","User ID"],nrows=21)

# exporting the given input-2 in the variable(i.e input2)
# before that while reading the file adding some extra paratmeter becuase different
# information present in single sheet and we only need the input table
input2 = pd.read_excel(xls,"(Input) Rigorbuilder RAW",skiprows=7,usecols=["name","uid","total_statements","total_reasons"],nrows=21)

'''to find the teamwise the rank is decided on the basis of sum of average statements and reasons
we need to first join the both input variable to find it, because the team name is present on input 1 and
other necessary feature/columns present in other input varibale.
So, we need to join the table using inner join function to proceed further.
'''
data = input1.merge(input2,left_on=["Name","User ID"],right_on=["name","uid"])[["User ID","Name","Team Name",
                                                                         "total_statements","total_reasons"]]

# finding the avg. statement and reason while aggregrating using group by function
aggregate_function = {
    "total_statements":"mean",
    "total_reasons": "mean"
}
leaderboard = data.groupby(["Team Name"],as_index=False).agg(aggregate_function).sort_values(by=["total_statements","total_reasons"],ascending=False)

# reseting the index because  sorting the values process we done above
leaderboard = leaderboard.reset_index(drop=True)

#fixing the naming convention erorr in the Team Name 'BrandTech Lab'
leaderboard.loc[8,"Team Name"] = "BrandTech Lab"

# again grouping by for merging both value of BrandTech Lab which name problem we showed just now
leaderboard = leaderboard.groupby(["Team Name"],as_index=False).agg(aggregate_function).sort_values(by=["total_statements",
                                                                                         "total_reasons"],
                                                                                     ascending=False).reset_index(drop=True)
leaderboard.index+=1


# Rounding the values to 2 decimal form
leaderboard["total_statements"] = leaderboard.total_statements.round(2)
leaderboard["total_reasons"] = leaderboard.total_reasons.round(2)

# creating a new column of rank using index value
leaderboard["Team Rank"] = leaderboard.index

# renaming the columns according to the given output(which is present in assignment)
leaderboard = leaderboard.rename(columns={"Team Name":"Thinking Teams Leaderboard",
                            "total_statements":"Average Statements",
                           "total_reasons":"Average Reasons"})
leaderboard = leaderboard[["Team Rank","Thinking Teams Leaderboard","Average Statements","Average Reasons"]]

# Output 2 code.
# Calculating each member of the rank according to no. of statemenet and reason
individual = input2.sort_values(by=["total_statements","total_reasons","name"],ascending=False)

# reseting the index because  sorting the values process we done above
individual = individual.reset_index(drop=True)
individual["Rank"] = individual.index+1

# renaming the columns according to the given output(which is present in assignment)
individual = individual.rename(columns={
    "name":"Name",
    "uid":"UID",
    "total_statements":"No. of Statements",
    "total_reasons": "No. of Reasons"
})
individual = individual[["Rank","Name","UID","No. of Statements","No. of Reasons"]]

# export the result in excel sheet
with pd.ExcelWriter('output.xlsx') as result:
    leaderboard.to_excel(result, sheet_name = 'Teamwise leaderboard', index = False)
    individual.to_excel(result, sheet_name = 'Individual leaderboard', index = False)


print('Successfully,Done!')
