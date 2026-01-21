import pandas as pd
from datetime import datetime, timedelta
from ortools.sat.python import cp_model
from collections import defaultdict
import os

class VolunteerScheduler:
    def __init__(self, input_file):
        self.input_file = input_file
        self.volunteers = []
        self.schedule_dates = []
        self.num_cameras = 7
        
    def load_volunteers(self):
        df = pd.read_excel(self.input_file)
        df.columns = df.columns.str.strip()
        
        for _, row in df.iterrows():
            is_priority = False
            if 'Priority' in row:
                val = row['Priority']
                if pd.notna(val) and (val is True or str(val).lower() in ['true', '1', 'yes']):
                    is_priority = True

            volunteer = {
                'name': str(row['Name']).strip(),
                'team': str(row['Team']).lower().strip(),
                'priority': is_priority,
                'preferred_day': row.get('Preferred_Day', None) if pd.notna(row.get('Preferred_Day', None)) else None,
                'frequency': str(row['Frequency']).strip().capitalize(),
                'allowed_date': pd.to_datetime(row['Allowed_Date']) if pd.notna(row['Allowed_Date']) else None,
                'unavailable_dates': self._parse_unavailable_dates(row.get('Unavailable_Dates', '')),
                'attached_person': row.get('Attached_Person', None) if pd.notna(row.get('Attached_Person', None)) else None,
                'camera_prefs': {
                    i+1: bool(row.get(f'Cam{i+1}_Pref', False)) 
                    for i in range(self.num_cameras)
                }
            }
            self.volunteers.append(volunteer)
        print(f"Loaded {len(self.volunteers)} volunteers.")
        return self.volunteers
    
    def _parse_unavailable_dates(self, date_str):
        if pd.isna(date_str) or not date_str: return []
        dates = []
        for d in str(date_str).split(','):
            try: dates.append(pd.to_datetime(d.strip()))
            except: pass
        return dates
    
    def generate_schedule_dates(self, target_month=None, target_year=None):
        """
        Generates dates for the specified month. 
        If no month/year provided, defaults to NEXT calendar month.
        """
        today = datetime.now()
        
        if target_month is None:
            # If current month is Dec (12), next is Jan (1)
            target_month = 1 if today.month == 12 else today.month + 1
            
        if target_year is None:
            # If we rolled over to Jan, increment the year
            target_year = today.year + 1 if (today.month == 12 and target_month == 1) else today.year

        print(f"Targeting Schedule for: {target_month}/{target_year}")

        first_day = datetime(target_year, target_month, 1)
        # Find the last day of the month
        if target_month == 12:
            last_day = datetime(target_year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(target_year, target_month + 1, 1) - timedelta(days=1)
            
        current = first_day
        dates = []
        while current <= last_day:
            if current.weekday() == 5: dates.append(('Saturday', current))
            elif current.weekday() == 6: dates.append(('Sunday', current))
            current += timedelta(days=1)
        
        self.schedule_dates = dates
        return dates
    
    def is_volunteer_available(self, volunteer, day_type, date):
        if volunteer['team'] == 'sub': return False
        if volunteer['team'] == 'saturday' and day_type != 'Saturday': return False
        if volunteer['team'] == 'sunday' and day_type != 'Sunday': return False
        for unavail_date in volunteer['unavailable_dates']:
            if date.date() == unavail_date.date(): return False
        
        freq = volunteer['frequency']
        allowed_date = volunteer['allowed_date']
        if freq in ['Default', 'Monthly'] and allowed_date:
            weeks_diff = (date - allowed_date).days // 7
            if freq == 'Default' and weeks_diff % 2 != 0: return False
            if freq == 'Monthly' and weeks_diff % 4 != 0: return False
        return True

    def solve_schedule(self):
        model = cp_model.CpModel()
        V, D, C = len(self.volunteers), len(self.schedule_dates), self.num_cameras

        assignments = {(v, d, c): model.NewBoolVar(f"a_v{v}_d{d}_c{c}") for v in range(V) for d in range(D) for c in range(1, C+1)}
        unfilled = {(d, c): model.NewBoolVar(f"u_d{d}_c{c}") for d in range(D) for c in range(1, C+1)}

        for d in range(D):
            for c in range(1, C+1):
                model.Add(sum(assignments[(v, d, c)] for v in range(V)) + unfilled[(d, c)] == 1)
        for v in range(V):
            for d in range(D):
                model.Add(sum(assignments[(v, d, c)] for c in range(1, C+1)) <= 1)

        for v, vol in enumerate(self.volunteers):
            has_any_pref = any(vol['camera_prefs'].values())
            for d, (day_type, date) in enumerate(self.schedule_dates):
                is_avail = self.is_volunteer_available(vol, day_type, date)
                for c in range(1, C+1):
                    if not is_avail:
                        model.Add(assignments[(v, d, c)] == 0)
                    if has_any_pref and not vol['camera_prefs'][c]:
                        model.Add(assignments[(v, d, c)] == 0)

        for v, vol in enumerate(self.volunteers):
            load = sum(assignments[(v, d, c)] for d in range(D) for c in range(1, C+1))
            freq = vol['frequency']
            max_cap = 1 if freq == 'Monthly' else (2 if freq == 'Default' else 4)
            model.Add(load <= max_cap)

        obj = []
        for d in range(D):
            for c in range(1, C+1):
                obj.append(-1000000 * unfilled[(d, c)])

        for v, vol in enumerate(self.volunteers):
            tier_weight = 100 if (vol['team'] in ['saturday', 'sunday'] and vol['frequency'] == 'Default') else 50
            if vol['priority']: tier_weight *= 50
            for d in range(D):
                for c in range(1, C+1):
                    obj.append(tier_weight * assignments[(v, d, c)])
                    if vol['camera_prefs'][c]:
                        obj.append(5000 * assignments[(v, d, c)])

        model.Maximize(sum(obj))
        solver = cp_model.CpSolver()
        if solver.Solve(model) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return self._extract_solution(solver, assignments, unfilled)
        return None

    def _extract_solution(self, solver, assignments, unfilled_positions):
        schedule = []
        volunteer_counts = defaultdict(int)
        total_unfilled = 0
        scheduled_names = set()
        
        for d_idx, (day_type, date) in enumerate(self.schedule_dates):
            day_schedule = {'Day': day_type, 'Date': date.strftime('%Y-%m-%d')}
            for camera in range(1, self.num_cameras + 1):
                if solver.Value(unfilled_positions[(d_idx, camera)]) == 1:
                    day_schedule[f'Camera {camera}'] = '** UNFILLED **'
                    total_unfilled += 1
                else:
                    for v_idx, volunteer in enumerate(self.volunteers):
                        if solver.Value(assignments[(v_idx, d_idx, camera)]) == 1:
                            name = volunteer['name']
                            day_schedule[f'Camera {camera}'] = name
                            volunteer_counts[name] += 1
                            scheduled_names.add(name)
                            break
            schedule.append(day_schedule)
        
        print("\n" + "="*50)
        print(f"{'SCHEDULING SUMMARY':^50}")
        print("="*50)

        print("\n[1] ASSIGNED VOLUNTEERS")
        sorted_counts = sorted(volunteer_counts.items(), key=lambda x: x[1], reverse=True)
        for name, count in sorted_counts:
            vol_info = next(v for v in self.volunteers if v['name'] == name)
            print(f" - {name:18} | Shifts: {count} | Team: {vol_info['team']:8} | Priority: {vol_info['priority']}")

        print("\n[2] NOT SCHEDULED")
        not_scheduled = [v for v in self.volunteers if v['name'] not in scheduled_names]
        for v in not_scheduled:
            reason = "(Sub Team)" if v['team'] == 'sub' else "(Not needed or constraint conflict)"
            print(f" - {v['name']:18} | Team: {v['team']:8} | {reason}")

        print("\n[3] COVERAGE")
        print(f" ✅ All slots filled" if total_unfilled == 0 else f" ⚠️ {total_unfilled} UNFILLED SLOTS")
        print("="*50 + "\n")
        
        return schedule

def main():
    # File configuration
    input_file = 'volunteers.xlsx'
    output_file = 'schedule_output.xlsx'
    
    scheduler = VolunteerScheduler(input_file)
    scheduler.load_volunteers()
    
    # ---------------------------------------------------------
    # DATE LOGIC:
    # To schedule NEXT month automatically, leave parameters empty:
    # scheduler.generate_schedule_dates()
    #
    # To override and schedule a SPECIFIC month:
    # scheduler.generate_schedule_dates(target_month=3, target_year=2026)
    # ---------------------------------------------------------
    
    scheduler.generate_schedule_dates() # Default: Next Calendar Month
    
    res = scheduler.solve_schedule()
    
    if res:
        df = pd.DataFrame(res)
        df.to_excel(output_file, index=False)
        print(f"Schedule exported to {output_file}")
    else:
        print("Error: Could not find a valid schedule.")

if __name__ == '__main__':
    main()