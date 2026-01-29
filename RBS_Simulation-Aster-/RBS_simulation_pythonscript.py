#import 
import xlwings as xw 
import os
import numpy as np 
import matplotlib.pyplot as plt

wb=xw.Book.caller()
ws=wb.sheets["RBS_Simulation"]

# --- Define states ---
Operational = "Operational"
Repair_immediate = "Repair_immediate"
Lead_time = "Lead_time"
Repair = "Repair"

#np.random.seed(42) #set seed for reproducibility
def run_rbs_simulation(qty, mbtf, mttr, lead_time, Simulation_Time):
    
    #list to track metrics
    timeline=[0] #timepoints, array contain 0 (start time), eventual timepoints will be appended here
    #downtime_list=[0] #downtime at each timepoint, array contain 0 (start time), eventual downtimes will be appended here
    status=[Operational] #status at each timepoint, array contain 'Operational' (start time), eventual statuses will be appended here
    qty_list=[qty] #spare quantity at each timepoint, array contain initial qty (start time), eventual quantities will be appended here

    #initialise simulation variables
    current_time=0 #start time
    next_time=0 #time of next event
    failure_time=0 #time to next failure (random generated from exponetial distribution)
    repair_time=0 #time to repair (random generated from exponetial distribution)
    downtime_with_spares=0 #total downtime when spares are available
    downtime_without_spares=0 #total downtime when no spares are available
    total_downtime=0 #total downtime
    uptime=0 #total time without failures
    failures_with_spares=0 #count of failures when spares are available
    failures_without_spares=0 #count of failures when no spares are available
   
#simulation loop 
#failure_time=np.random.exponential(mbtf) #generate time to next failure, this is a random variable where the average is(mbtf)
#repair_time=np.random.exponential(mttr) #generate repair time
    while current_time<=Simulation_Time: #while we are within the simulation time
        failure_time=np.random.exponential(mbtf) #generate time to next failure
        next_time=current_time+failure_time #calculate next event time
        #current_time+=failure_time #advance time by failure time
        if next_time>=Simulation_Time:
            # Ensure the simulation timeline always ends exactly at Simulation_Time
            if timeline[-1] < Simulation_Time:
                timeline.append(Simulation_Time)
                # Extend the last status to the end
                status.append(Repair_immediate if qty > 0 else Repair)
                qty-=1
                qty_list.append(qty)  # last known quantity
            break
        current_time+=failure_time #advance time by failure time
        timeline.append(current_time) #append current time to timeline
        if qty>0: #if there are spare parts
            failures_with_spares+=1 #increment failures with spares
            status.append(Repair_immediate) #append status
            qty-=1 #reduce spare parts by 1
            #timeline.append(current_time) #append current time to timeline
            qty_list.append(qty) #append current qty to qty_list
            
            repair_time=np.random.exponential(mttr) #generate repair time
            downtime_with_spares+=repair_time #add repair time to downtime with spares
            next_time+=repair_time #advance time by repair time
            if next_time>Simulation_Time: #if we have exceeded simulation time
                if timeline[-1]<Simulation_Time:
                    timeline.append(Simulation_Time)
                    status.append(Operational)
                    qty_list.append(qty)
            current_time+=repair_time #advance time by repair time
            timeline.append(current_time) #append current time to timeline
            #status.append(Repair_immediate) #append status
            #qty-=1 #reduce spare parts by 1
            #timeline.append(current_time) #append current time to timeline
            status.append(Operational) #append status
            qty_list.append(qty) #append current qty to qty_list
        else: #if no spare parts
            failures_without_spares+=1 #increment failures without spares
            #timeline.append(current_time) #append current time to timeline
            status.append(Lead_time) #append status
            qty_list.append(qty) #append current qty to qty_list
            #if current_time>Simulation_Time: #if we have exceeded simulation time
                #break #exit loop
            downtime_without_spares+=lead_time #add lead time
            next_time+=lead_time #advance time by lead time
            if next_time>Simulation_Time: #if we have exceeded simulation time
                if timeline[-1]<Simulation_Time:
                    timeline.append(Simulation_Time)
                    status.append(Repair)
                    qty_list.append(qty)
            current_time+=lead_time #advance time by lead time
            timeline.append(current_time) #append current time to timeline
            status.append(Repair) #append status
            qty_list.append(qty) #append current qty to qty_list
            #if current_time>Simulation_Time: #if we have exceeded simulation time
                #break #exit loop
            repair_time=np.random.exponential(mttr) #generate repair time
            downtime_without_spares+=repair_time #add repair time to downtime without spares
            next_time+=repair_time #advance time by repair time
            if next_time>Simulation_Time: #if we have exceeded simulation time
                if timeline[-1]<Simulation_Time:
                    timeline.append(Simulation_Time)
                    status.append(Operational)
                    qty_list.append(qty)
            current_time+=repair_time #advance time by repair time
            timeline.append(current_time) #append current time to timeline
            status.append(Operational) #append status
            qty_list.append(qty) #append current    qty to qty_list
            #if current_time>Simulation_Time: #if we have exceeded simulation time
                #break #exit loop
                #qty remains 0 as no spares are available
                
    total_downtime=downtime_with_spares+downtime_without_spares #calculate total downtime
    uptime=current_time-total_downtime #calculate time without failures
    availability = uptime / current_time if current_time > 0 else 0.0 #calculate availability
    #ideally, we cant forsee damand(failure) so we determine fill rate instead, where every simulation will be classified to either availability meet demand or not
    #we then calculate the amount of times it meet demand over total simulations. This gives us P(K<=s), P(Demand<=spare)
    fill_rate = failures_with_spares / (failures_with_spares + failures_without_spares) if (failures_with_spares + failures_without_spares) > 0 else 0.0
    stockout_probability = 1 - fill_rate
    readiness = fill_rate*availability
    return total_downtime, uptime, downtime_without_spares, downtime_with_spares, availability, timeline, status, qty_list, fill_rate, stockout_probability, readiness

#visualization
def plot_simple_gantt(timeline, status, qty):

    STATE_COLORS = {
        Operational: "green",
        Repair_immediate: "blue",
        Lead_time: "red",
        Repair: "orange"
    }   
    fig, ax = plt.subplots(figsize=(12,2))
    y = 0 #single bar at y=0
    height = 0.4

    for i in range(len(status)-1):
        start = timeline[i]
        end = timeline[i+1] #if i+1 < len(timeline) else timeline[i] #end of current segment
        ax.barh(y, end-start, left=start, height=height,
                color=STATE_COLORS.get(status[i], "gray"), edgecolor="black")
    ax.set_yticks([y])
    ax.set_yticklabels(["System"])
    ax.set_xlabel("Time")
    ax.set_title("System State Gantt Chart")

    # Legend
    handles = [plt.Rectangle((0,0),1,1,color=color) for color in STATE_COLORS.values()]
    labels = STATE_COLORS.keys()
    ax.legend(handles, labels, loc="upper center", ncol=4)
    plt.tight_layout()
    
    #save figure as image
    img_path = os.path.join(os.getcwd(), f"Gantt_Qty{qty}.png")
    fig.savefig(img_path)
    plt.close(fig)

    # Insert image into Excel
    col_start = 'L'  # Starting column for images
    row= 6 + 6*(qty_values.index(qty))  # Row based on quantity index
    ws.pictures.add(img_path, name=f"Gantt_Qty{qty}", update=True,
                    left=ws.range(f'{col_start}{row}').left,
                    top=ws.range(f'{col_start}{row}').top,
                    width=500, height=80)
    

#monte carlo simulation for each spare quantity value
def multiple_simulations_per_qty(qty_values, num_simulations):
    total_downtime_results=[]
    uptime_results=[]
    downtime_without_spares_results=[]
    downtime_with_spares_results=[]
    availability_results=[]
    #timeline_results=[]
    #status_results=[]
    #qty_list_results=[]
    fill_rate_results=[]
    stockout_probability_results=[]
    readiness_results=[]
    for row,q in enumerate(qty_values,start=2): #start from row 2 in excel
        total_downtime, uptime, downtime_without_spares, downtime_with_spares, availability, timeline, status, qty_list, fill_rate, stockout_probability, readiness = run_rbs_simulation(q, mbtf, mttr, lead_time, Simulation_Time)
        plot_simple_gantt(timeline, status,q) #plot gantt chart for each spare quantity
        # #metrics 
        # print("=== Simulation Outputs ===")
        # print(f"Total time machine was operational: {uptime:.2f}")
        # print(f"Total downtime when spares were 0: {downtime_without_spares:.2f}")
        # print(f"Total downtime with spares: {downtime_with_spares:.2f}")
        # print(f"Overall availability: {availability:.2%}")
        # #write outputs to excel
        # ws.range(f'G{row}').value = uptime
        # ws.range(f'H{row}').value = downtime_without_spares
        # ws.range(f'I{row}').value = downtime_with_spares
        # ws.range(f'J{row}').value = availability 
        
        total_downtime_avg=[]
        uptime_avg=[]
        downtime_without_spares_avg=[]
        downtime_with_spares_avg=[]
        availability_avg=[]
            #timeline_avg=[]
            #status_avg=[]
            #qty_list_avg=[]
        fill_rate_avg=[]
        stockout_probability_avg=[]
        readiness_avg=[]
            
        for i in range(num_simulations):
            total_downtime, uptime, downtime_without_spares, downtime_with_spares, availability, timeline, status, qty_list, fill_rate, stockout_probability, readiness = run_rbs_simulation(q, mbtf, mttr, lead_time, Simulation_Time)
            
            total_downtime_avg.append(total_downtime)
            uptime_avg.append(uptime)
            downtime_without_spares_avg.append(downtime_without_spares)
            downtime_with_spares_avg.append(downtime_with_spares)
            availability_avg.append(availability)
            fill_rate_avg.append(fill_rate)
            stockout_probability_avg.append(stockout_probability)
            readiness_avg.append(readiness)
         #metrics 
        print("=== Simulation Outputs ===")
        print(f"Average time machine was operational: {np.mean(uptime_avg):.2f}")
        print(f"Average downtime when spares were 0: {np.mean(downtime_without_spares_avg):.2f}")
        print(f"Average downtime with spares: {np.mean(downtime_with_spares_avg):.2f}")
        print(f"Average availability: {np.mean(availability_avg):.2%}")
        #write outputs to excel
        ws.range(f'H1').value = "Avg Uptime"
        ws.range(f'H{row}').value = np.mean(uptime_avg)
        ws.range(f'I1').value = "Avg Downtime without Spares"
        ws.range(f'I{row}').value = np.mean(downtime_without_spares_avg)
        ws.range(f'J1').value = "Avg Downtime with Spares"
        ws.range(f'J{row}').value = np.mean(downtime_with_spares_avg)
        ws.range(f'K1').value = "Avg Availability"
        ws.range(f'K{row}').value = np.mean(availability_avg)
        
        
        total_downtime_results.append(np.mean(total_downtime_avg))
        uptime_results.append(np.mean(uptime_avg))
        downtime_without_spares_results.append(np.mean(downtime_without_spares_avg))
        downtime_with_spares_results.append(np.mean(downtime_with_spares_avg))
        availability_results.append(np.mean(availability_avg))
        fill_rate_results.append(np.mean(fill_rate_avg))
        stockout_probability_results.append(np.mean(stockout_probability_avg))
        readiness_results.append(np.mean(readiness_avg))
    print(total_downtime_results, uptime_results, downtime_without_spares_results, downtime_with_spares_results, availability_results, fill_rate_results, stockout_probability_results, readiness_results)
    return total_downtime_results, uptime_results, downtime_without_spares_results, downtime_with_spares_results, availability_results, fill_rate_results, stockout_probability_results, readiness_results
    



#read values from excel tables, simulation based on user inputs
qty=ws.range('B2').expand('down').value #read spare quantity from excel, expand down to read multiple values
if not isinstance(qty, list):
    qty = [qty]
qty_values=[q for q in qty if q is not None] #filter out None values
mbtf=int(ws.range('C2').value) 
mttr=int(ws.range('D2').value)
lead_time=int(ws.range('E2').value)
Simulation_Time=int(ws.range('F2').value)
num_simulations=int(ws.range('G2').value)

#call function to run multiple simulations per spare quantity
multiple_simulations_per_qty(qty_values, num_simulations)





# #function to run multiple simulations and average results   
# def run_multiple_simulations(num_simulations):
#     availability_results = []
#     for q in qty_values:
#         for i in range(num_simulations):
#             _, _, _, _, availability, _, _, _, _, _ = run_rbs_simulation(q, mbtf, mttr, lead_time, Simulation_Time)
#             availability_results.append(availability)
#             average_availability = np.mean(availability_results)
#     return average_availability







# plt.figure(figsize=(12, 4))
# plt.step(timeline_user_in, status_user_in, where='post', linewidth=2)
# plt.yticks([0, 1, 2, immediate = "Repair_immediate", "Operational", "Failure_without_spare", "Repair_no_spare"])
# plt.xlabel("Time")
# plt.ylabel("Machine Status")
# plt.title("RBS Simulation Timeline")
# plt.grid(True)
# plt.show()


# #simulation based on an array of available spare parts
# spare_parts_array=[0, 1, 2, 3, 4]
# availability_results=[]


# for spare_qty in spare_parts_array:
#     uptime, downtime_without_spares, downtime_with_spares, availability = run_rbs_simulation(spare_qty,(mbtf),(mttr), lead_time, Simulation_Time)
#     availability_results.append(availability)

# # Plot availability vs spare parts
# plt.figure(figsize=(10, 6))
# plt.plot(spare_parts_array, availability_results, marker='o')
# plt.xlabel("Number of Spare Parts")
# plt.ylabel("Availability")
# plt.title("Availability vs Number of Spare Parts")
# plt.grid(True)
# plt.show()