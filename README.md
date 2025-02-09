#  **Optimal Surgery Scheduling System**

##  **Project Description**

The objective of this project is to develop a **surgery scheduling system** that organizes the list of patients to be operated on based on two key factors:

1. **Urgency level** of the surgery.
2. **Availability** of operating rooms and surgeons.

The system generates a schedule that includes:

- **Patient ID** to be operated.
- **Assigned surgeon**.
- **Operating room** where the surgery will take place.
- **Scheduled date** for the surgery.

---

## **Problem Context**

The problem addresses the complexity of coordinating limited medical resources, such as operating rooms and surgeons with specific **skills**, while prioritizing the most critical patients. However, resource constraints may lead to scenarios where less critical patients are scheduled before more urgent ones.

### **Identified Challenges:**
- **Limited number of operating rooms** available for specific types of surgeries.
- **Specialization of medical staff**, as not all surgeons can perform every type of surgery.
- **Efficient management of available time slots** in the surgical schedule.

---

## **Results and Conclusions**

The developed model **prioritizes the most critical patients**, but logistical constraints can lead to scenarios where:

- **Less critical patients** are scheduled before more urgent ones due to the specific availability of operating rooms and surgeons.
- **Available slots** are filled with the first patient meeting the necessary conditions (suitable operating room and available surgeon).

This approach allows for the **optimal use of resources** without leaving idle times in the schedule but may require adjustments to improve **medical prioritization**.

---

## **Technologies Used**

- **Python** for algorithm development.
- **Pandas** for data management and analysis.
- **Google OR-Tools** for resource optimization.
- **Matplotlib/Seaborn** for result visualization.

---

## **Installation**

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/surgery_scheduling.git
   cd surgery_scheduling
   ```

2. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the program:
   ```bash
   python main.py
   ```

---

## **Project Structure**

```
/
├── data/                  # Input data (patients, operating rooms, surgeons)
├── src/                   # Source code of the algorithm    
└── README.md              # Project documentation
```


---

**Version 1.0** - Automated surgery scheduling project.

