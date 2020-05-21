#!/usr/bin/env python
# coding: utf-8

# In[ ]:


def optimize(inputFile,outputFile):
    """
    Description:
           The function for do class scheduling, which can 
           maximize overall departments' preference scores,
           meanwhile promoting fairness.
           
    Arguments
        1. InputFile: the path contains the Excel file in which preferences data is saved for optimization
        2. OutputFile: the path customized and defined by users to produce the optimal class schedule
    
    """
    
    import pandas as pd
    from gurobipy import Model, GRB
    
    # Classroom settings - R
    classroom = pd.read_excel(inputFile,sheet_name='classroom - Pr',index_col = 0)

    # The number of hours defined for each timeslot - ht
    hours = pd.read_excel(inputFile,sheet_name='hours - ht',index_col = 0)

    # How many hours to teach in a week
    md =  pd.read_excel(inputFile,sheet_name='md',index_col = 0)

    # Departments Preference Scores
    department_dict = {i:pd.read_excel(inputFile,sheet_name=i,index_col = 0)for i in md.index}

    # Set of department
    department = md.index

    # Set of weekday
    weekday = department_dict['BUCO'].columns

    # Set of timeslots
    timeslots = department_dict['BUCO'].index[1:]

    # Set of classrooms
    classroom_set = classroom.index

    # Whether a classroom is a big or small classrooms
    br = classroom.Big

    # hour to teach
    ht = md['hours to teach in a semester']

    # Initialize the Models
    mod = Model()

    # Add variables - a{dtrc}
    X = mod.addVars(department,timeslots,classroom_set,weekday,lb = 0,vtype = GRB.BINARY)
    U = mod.addVar(lb = 0)
    L = mod.addVar(lb = 0)

    ## Objective Function
    mod.setObjective(sum(X[d,t,r,c]*department_dict[d].loc[t,c] for d in department for t in timeslots for c in weekday for r in classroom_set) - (U - L),                     sense=GRB.MAXIMIZE
    )

    ## Constraints - MW,TH
    for d in department:
        for t in timeslots[:-1]:
            for c in weekday[:2]:
                for r in classroom_set:
                    mod.addConstr(X[d,t,r,c] == X[d,t,r,c+2], name =f'MWTH_{d}{t}{r}{c}')

    ## Constraints - full arrangement 
    for t in timeslots:
        for c in weekday:
            for r in classroom_set:
                mod.addConstr(sum(X[d,t,r,c] for d in department) <=1, name = f'full_arrange_{d}_{t}_{r}_{c}')

    ## Constraint - At least one big room for each department
    for d in department:
        mod.addConstr(sum(X[d,t,r,c] * br[r] for t in timeslots for r in classroom_set for c in weekday )>=1, name = f'full_arrange_{d}' )

    ## Constraint - Preference Balance -  Lower Bound
    for d in department:
        nominator = sum(X[d,t,r,c]*department_dict[d].loc[t,c] for t in timeslots for r in classroom_set for c in weekday)

        denominator = sum(department_dict[d].loc[t,c] for t in timeslots for r in classroom_set)

        mod.addConstr((nominator/denominator) >= L,name = f'Preference_Bal_L_{d}')

    ## Constraint - Preference Balance -  Upper Bound    
    for d in department:
        nominator = sum(X[d,t,r,c]*department_dict[d].loc[t,c] for t in timeslots for r in classroom_set for c in weekday)

        denominator = sum(department_dict[d].loc[t,c] for t in timeslots for r in classroom_set)

        mod.addConstr((nominator/denominator) <= U,name = f'Preference_Bal_U_{d}')

    ## Constraint - hour demands
    for d in department:
        mod.addConstr(sum(X[d,t,r,c]*hours.loc[t,c] for t in timeslots for r in classroom_set for c in weekday )>=ht[d], name = f'hour_demand_{d}')

    # build up a storage dataframe
    index_1 = pd.MultiIndex.from_product([weekday,timeslots],
                                        names=['Weekday ','Timeslot']
                                       )
    # Create an empty dataframe to store the optimized outputs from the model
    schedule = pd.DataFrame(index = index_1, 
                            columns = classroom_set
    )

    # Functionality to perform optimization and print out the results
    mod.setParam('OutputFlag',False)
    mod.optimize()
    print(f'The Optimal Preference Scores is: {mod.ObjVal}')

    # Get the output
    for d in department:
        for t in timeslots:
            for c in weekday:
                for r in classroom_set:
                    # only select the value of a variable equal to 1.0
                    if X[d,t,r,c].x:
                        schedule.loc[(c,t),r] = d
    
    # to get a clean, organized and concise timeslot timetable
    df_storage = []
    idx = pd.IndexSlice
    for c in weekday:
        if c in [1,2]:
            df_temp = schedule.loc[idx[c,timeslots[:-1]],:].reset_index()
            if c == 1:
                df_temp.loc[:,'Weekday '] = 'Mon & Wed'
            elif c == 2:
                df_temp.loc[:,'Weekday '] = 'Tue & Thur'
            df_storage.append(df_temp) 

        if c in [1,2,3,4]:
            df_temp = schedule.loc[idx[c,timeslots[-1]],:].to_frame().T.reset_index()

            df_temp.columns = ['Weekday ', 'Timeslot', 102, 104, 110, 112, 202, 204, 210, 212]
            if c == 1:
                df_temp.loc[:,'Weekday '] = 'Mon'
            elif c == 2:
                df_temp.loc[:,'Weekday '] = 'Tue'
            elif c == 3:
                df_temp.loc[:,'Weekday '] = 'Wed'
            elif c == 4:
                df_temp.loc[:,'Weekday '] = 'Thur'     
            df_storage.append(df_temp)
        if c in [5]:
            df_temp = schedule.loc[idx[c,timeslots[:]],:].reset_index()
            df_temp.loc[:,'Weekday '] = 'Fri'
            df_storage.append(df_temp)

    # to sort the table in the order of ["Mon & Wed", "Mon", "Wed", "Tue & Thur", "Tue", "Thur", "Fri"]
    df_schedule = pd.concat(df_storage, axis = 0)      
    orders = ["Mon & Wed", "Mon", "Wed", "Tue & Thur", "Tue", "Thur", "Fri"]
    idx_ordered = pd.Categorical(df_schedule['Weekday '].values, categories=orders,
              ordered=True)
    
    df_schedule['Weekday '] = idx_ordered
    
    # define the timeslot order
    orders2 = timeslots
    idx_ordered2 = pd.Categorical(df_schedule['Timeslot'].values, categories=orders2,
              ordered=True)
    df_schedule['Timeslot'] = idx_ordered2
    
    # sort values by 'Weekday ','Timeslot'
    df_schedule.sort_values(by = ['Weekday ','Timeslot'],ascending=True,inplace = True)
    df_schedule.set_index(['Weekday ','Timeslot'],inplace = True)
    
    # Create an output file
    writer = pd.ExcelWriter(outputFile)
    df_schedule.to_excel(writer, sheet_name = 'Class Schedule', index = True)
    writer.save()

import os, sys
    
if __name__ == "__main__":
    if len(sys.argv) !=3:
    # python book.py inputfile.xlsx, outputfile.xlsx <- 3 argument
        print('Not Correct Syntax')

    else:
        inputFile = sys.argv[1]
        outputFile = sys.argv[2]
        
    if os.path.exists(inputFile):
        optimize(inputFile, outputFile)
        print('successfully optimize')

    else:
        print('input file: {} not found'.format(inputfile))

