import os
import pandas as pd
from tabulate import tabulate

# Fungsi untuk membaca data dari Excel
def read_process_data(file_path):
    df = pd.read_excel(file_path)
    processes = df.to_dict(orient="records")  # Convert to list of dictionaries
    return processes

def write_to_excel_no_3_4(data, filename):
    # Tentukan nama kolom untuk no 3 dan 4
    columns = ["PID", "Arrival Time", "Burst Time", "Remaining Time", "Waiting Time", "Turnaround Time", "Current Time"]

    # Membuat DataFrame menggunakan data dan kolom yang sesuai
    df = pd.DataFrame(data, columns=columns)

    # Pastikan folder 'output' ada
    os.makedirs("output", exist_ok=True)

    # Menyimpan DataFrame ke file Excel
    file_path = f"output/{filename}.xlsx"
    df.to_excel(file_path, index=False)
    print(f"File berhasil disimpan di {file_path}")

# Fungsi untuk menulis output ke Excel
def write_to_excel(data, algorithm_name):
    # Pastikan folder output ada, jika tidak maka buat foldernya
    os.makedirs("output", exist_ok=True)
    
    # Simpan file di dalam folder output
    file_name = f"output/{algorithm_name}_scheduling_output.xlsx"
    df = pd.DataFrame(data, columns=["PID", "Arrival Time", "Burst Time", "Waiting Time", "Turnaround Time"])
    df.to_excel(file_name, index=False)
    print(f"Hasil {algorithm_name} telah disimpan ke file {file_name}.\n")

# Fungsi untuk First Come First Serve (FCFS)
def fcfs_scheduling(processes):
    processes.sort(key=lambda x: x["arrival_time"])  # Sort by arrival time
    current_time = 0
    waiting_times = []
    turnaround_times = []
    results = []

    for process in processes:
        if current_time < process["arrival_time"]:
            current_time = process["arrival_time"]
        waiting_time = current_time - process["arrival_time"]
        turnaround_time = waiting_time + process["burst_time"]
        waiting_times.append(waiting_time)
        turnaround_times.append(turnaround_time)
        current_time += process["burst_time"]
        
        # Tambahkan data proses ke dalam list
        results.append([process["PID"], process["arrival_time"], process["burst_time"], waiting_time, turnaround_time])

    # Tampilkan tabel
    print("FCFS Scheduling:")
    print(tabulate(results, headers=["PID", "Arrival Time", "Burst Time", "Waiting Time", "Turnaround Time"], tablefmt="grid"))

    # Simpan ke Excel
    write_to_excel(results, "FCFS")

    avg_waiting_time = sum(waiting_times) / len(waiting_times)
    avg_turnaround_time = sum(turnaround_times) / len(turnaround_times)
    return avg_waiting_time, avg_turnaround_time

# Fungsi untuk Shortest Job First (Non-Preemptive)
def sjf_non_preemptive(processes):
    processes.sort(key=lambda x: (x["arrival_time"], x["burst_time"]))
    current_time = 0
    waiting_times = []
    turnaround_times = []
    remaining_processes = processes.copy()
    results = []

    while remaining_processes:
        available_processes = [p for p in remaining_processes if p["arrival_time"] <= current_time]
        if available_processes:
            process = min(available_processes, key=lambda x: x["burst_time"])
            waiting_time = current_time - process["arrival_time"]
            turnaround_time = waiting_time + process["burst_time"]
            waiting_times.append(waiting_time)
            turnaround_times.append(turnaround_time)
            current_time += process["burst_time"]
            remaining_processes.remove(process)

            # Tambahkan data proses ke dalam list
            results.append([process["PID"], process["arrival_time"], process["burst_time"], waiting_time, turnaround_time])
        else:
            current_time += 1  # Increment time if no process is available

    # Tampilkan tabel
    print("SJF (Non-Preemptive) Scheduling:")
    print(tabulate(results, headers=["PID", "Arrival Time", "Burst Time", "Waiting Time", "Turnaround Time"], tablefmt="grid"))

    # Simpan ke Excel
    write_to_excel(results, "SJF_Non_Preemptive")

    avg_waiting_time = sum(waiting_times) / len(waiting_times)
    avg_turnaround_time = sum(turnaround_times) / len(turnaround_times)
    return avg_waiting_time, avg_turnaround_time

# Fungsi untuk Round Robin (Quantum Time = 12)
def round_robin(processes, quantum=12):
    remaining_times = {p["PID"]: p["burst_time"] for p in processes}
    current_time = 0
    waiting_times = {p["PID"]: 0 for p in processes}
    last_execution_time = {p["PID"]: p["arrival_time"] for p in processes}
    results = []

    while any(remaining_times.values()):
        for process in processes:
            pid = process["PID"]
            if remaining_times[pid] > 0 and process["arrival_time"] <= current_time:
                time_slice = min(quantum, remaining_times[pid])
                waiting_times[pid] += current_time - last_execution_time[pid]
                current_time += time_slice
                remaining_times[pid] -= time_slice
                last_execution_time[pid] = current_time

                # Tambahkan data proses ke dalam list (menggunakan remaining time untuk representasi)
                results.append([pid, process["arrival_time"], process["burst_time"], waiting_times[pid], current_time])

    # Tampilkan tabel
    print("Round Robin Scheduling (Quantum=12):")
    print(tabulate(results, headers=["PID", "Arrival Time", "Burst Time", "Waiting Time", "Current Time"], tablefmt="grid"))

    # Simpan ke Excel
    write_to_excel(results, "Round_Robin")

    avg_waiting_time = sum(waiting_times.values()) / len(waiting_times)
    avg_turnaround_time = avg_waiting_time + sum(p["burst_time"] for p in processes) / len(processes)
    return avg_waiting_time, avg_turnaround_time

# Fungsi untuk Longest Job First (Preemptive)
def ljf_preemptive(processes):
    processes = sorted(processes, key=lambda x: x["arrival_time"])
    n = len(processes)
    remaining_times = {p["PID"]: p["burst_time"] for p in processes}
    waiting_times = {p["PID"]: 0 for p in processes}
    turnaround_times = {p["PID"]: 0 for p in processes}
    
    current_time = 0
    completed = 0
    last_process = None
    results = []

    while completed < n:
        available_processes = [p for p in processes if p["arrival_time"] <= current_time and remaining_times[p["PID"]] > 0]
        if available_processes:
            process = max(available_processes, key=lambda x: remaining_times[x["PID"]])
            pid = process["PID"]

            if last_process and last_process != pid:
                waiting_times[pid] += current_time - process["arrival_time"]

            remaining_times[pid] -= 1
            current_time += 1

            if remaining_times[pid] == 0:
                completed += 1
                turnaround_times[pid] = current_time - process["arrival_time"]
                waiting_times[pid] = turnaround_times[pid] - process["burst_time"]

            last_process = pid
            results.append([pid, process["arrival_time"], process["burst_time"], remaining_times[pid], waiting_times[pid], turnaround_times[pid], current_time])
        else:
            current_time += 1

    avg_waiting_time = sum(waiting_times.values()) / n
    avg_turnaround_time = sum(turnaround_times.values()) / n

    # Simpan ke Excel
    write_to_excel_no_3_4(results, "LJF_preemptive")

    print("Longest Job First (Preemptive) Scheduling:")
    print(tabulate(results, headers=["PID", "Arrival Time", "Burst Time", "Remaining Time", "Waiting Time", "Turnaround Time", "Current Time"], tablefmt="grid"))
    return avg_waiting_time, avg_turnaround_time

# Fungsi untuk Shortest Job First (Preemptive)
def sjf_preemptive(processes):
    processes = sorted(processes, key=lambda x: x["arrival_time"])
    n = len(processes)
    remaining_times = {p["PID"]: p["burst_time"] for p in processes}
    waiting_times = {p["PID"]: 0 for p in processes}
    turnaround_times = {p["PID"]: 0 for p in processes}
    
    current_time = 0
    completed = 0
    last_process = None
    results = []

    while completed < n:
        available_processes = [p for p in processes if p["arrival_time"] <= current_time and remaining_times[p["PID"]] > 0]
        if available_processes:
            process = min(available_processes, key=lambda x: remaining_times[x["PID"]])
            pid = process["PID"]
            
            if last_process and last_process != pid:
                waiting_times[pid] += current_time - process["arrival_time"]

            remaining_times[pid] -= 1
            current_time += 1

            if remaining_times[pid] == 0:
                completed += 1
                turnaround_times[pid] = current_time - process["arrival_time"]
                waiting_times[pid] = turnaround_times[pid] - process["burst_time"]

            last_process = pid
            results.append([pid, process["arrival_time"], process["burst_time"], remaining_times[pid], waiting_times[pid], turnaround_times[pid], current_time])
        else:
            current_time += 1

    print("Shortest Job First (Preemptive) Scheduling:")
    print(tabulate(results, headers=["PID", "Arrival Time", "Burst Time", "Remaining Time", "Waiting Time", "Turnaround Time", "Current Time"], tablefmt="grid"))

    # Simpan ke Excel
    write_to_excel_no_3_4(results, "SJF_preemptive")

    avg_waiting_time = sum(waiting_times.values()) / n
    avg_turnaround_time = sum(turnaround_times.values()) / n

    return avg_waiting_time, avg_turnaround_time

# Main program
file_path = "processes.xlsx"  # Path to the Excel file
processes = read_process_data(file_path)

while True:
    print("=================")
    print("     SCHEDULE    ")
    print("=================")
    print("1. First Come First Serve")
    print("2. Shortest Job First (Non-Preemptive)")
    print("3. Shortest Job First (Preemptive)")
    print("4. Longest Job First (Preemptive)")
    print("5. Round Robin (Quantum time = 12)")
    print("6. Exit")

    choice = input("Please select an option (1-6): ")

    if choice == "1":
        print("\nFCFS Scheduling:")
        fcfs_avg_wait, fcfs_avg_turnaround = fcfs_scheduling(processes)
        print(f"Average Waiting Time: {fcfs_avg_wait}, Average Turnaround Time: {fcfs_avg_turnaround}\n")

    elif choice == "2":
        print("\nSJF (Non-Preemptive) Scheduling:")
        sjf_avg_wait, sjf_avg_turnaround = sjf_non_preemptive(processes)
        print(f"Average Waiting Time: {sjf_avg_wait}, Average Turnaround Time: {sjf_avg_turnaround}\n")

    elif choice == "3":
        print("\nSJF (Preemptive) Scheduling:")
        sjf_pre_avg_wait, sjf_pre_avg_turnaround = sjf_preemptive(processes)
        print(f"Average Waiting Time: {sjf_pre_avg_wait}, Average Turnaround Time: {sjf_pre_avg_turnaround}\n")

    elif choice == "4":
        print("\nLJF (Preemptive) Scheduling:")
        ljf_pre_avg_wait, ljf_pre_avg_turnaround = ljf_preemptive(processes)
        print(f"Average Waiting Time: {ljf_pre_avg_wait}, Average Turnaround Time: {ljf_pre_avg_turnaround}\n")

    elif choice == "5":
        print("\nRound Robin Scheduling (Quantum=12):")
        rr_avg_wait, rr_avg_turnaround = round_robin(processes)
        print(f"Average Waiting Time: {rr_avg_wait}, Average Turnaround Time: {rr_avg_turnaround}\n")

    elif choice == "6":
        print("Exiting the program.")
        break

    else:
        print("Invalid option. Please choose a number between 1 and 6.")

