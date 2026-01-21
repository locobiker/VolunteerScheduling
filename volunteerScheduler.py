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
                'frequency': str(row['Frequency']).strip(),
                'allowed_date': pd.to_datetime(row['Allowed_Date']) if pd.notna(row['Allowed_Date']) else None,
                'unavailable_dates': self._parse_unavailable_dates(row.get('Unavailable_Dates', '')),
                'attached_person': row.get('Attached_Person', None) if pd.notna(row.get('Attached_Person', None)) else None,
                'camera_prefs': {
                    i+1: bool(row.get(f'Cam{i+1}_Pref', False)) 
                    for i in range(self.num_cameras)
                }
            }
            self.volunteers.append(volunteer)
        return self.volunteers
    
    def _parse_unavailable_dates(self, date_str):
        if pd.isna(date_str) or not date_str: return []
        dates = []
        for d in str(date_str).split(','):
            try: dates.append(pd.to_datetime(d.strip()))
            except: pass
        return dates
    
    def generate_schedule_dates(self, target_month, target_year):
        first_day = datetime(target_year, target_month, 1)
        last_day = (datetime(target_year, target_month + 1, 1) if target_month < 12 else datetime(target_year + 1, 1, 1)) - timedelta(days=1)
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
        
        freq = str(volunteer['frequency']).strip().capitalize()
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

        # Coverage & Max 1 per day
        for d in range(D):
            for c in range(1, C+1):
                model.Add(sum(assignments[(v, d, c)] for v in range(V)) + unfilled[(d, c)] == 1)
        for v in range(V):
            for d in range(D):
                model.Add(sum(assignments[(v, d, c)] for c in range(1, C+1)) <= 1)

        # HARD CONSTRAINTS: Availability and Camera Training
        for v, vol in enumerate(self.volunteers):
            # Check if volunteer has ANY preferences selected
            has_any_pref = any(vol['camera_prefs'].values())
            
            for d, (day_type, date) in enumerate(self.schedule_dates):
                is_avail = self.is_volunteer_available(vol, day_type, date)
                for c in range(1, C+1):
                    # 1. Availability check
                    if not is_avail:
                        model.Add(assignments[(v, d, c)] == 0)
                    
                    # 2. STRICT CAMERA ENFORCEMENT: 
                    # If they marked specific cameras, they CANNOT be on any others.
                    if has_any_pref and not vol['camera_prefs'][c]:
                        model.Add(assignments[(v, d, c)] == 0)

        # Hard Caps for Frequency
        volunteer_load = [model.NewIntVar(0, D, f"load_{v}") for v in range(V)]
        for v, vol in enumerate(self.volunteers):
            model.Add(volunteer_load[v] == sum(assignments[(v, d, c)] for d in range(D) for c in range(1, C+1)))
            freq = str(vol['frequency']).strip().capitalize()
            max_cap = 1 if freq == 'Monthly' else (2 if freq == 'Default' else 4)
            model.Add(volunteer_load[v] <= max_cap)

        # Objective Function
        obj = []
        for d in range(D):
            for c in range(1, C+1):
                obj.append(-1000000 * unfilled[(d, c)]) # Fill positions first

        for v, vol in enumerate(self.volunteers):
            freq = str(vol['frequency']).strip().capitalize()
            # Tiered weighting
            tier_weight = 100 if (vol['team'] in ['saturday', 'sunday'] and freq == 'Default') else 50
            if vol['priority']: tier_weight *= 50
            
            for d in range(D):
                for c in range(1, C+1):
                    # Base priority assignment
                    obj.append(tier_weight * assignments[(v, d, c)])
                    # High bonus for preferred cameras
                    if vol['camera_prefs'][c]:
                        obj.append(5000 * assignments[(v, d, c)])

        model.Maximize(sum(obj))
        solver = cp_model.CpSolver()
        if solver.Solve(model) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return self._extract_solution(solver, assignments, unfilled)
        return None

    def _extract_solution(self, solver, assignments, unfilled_positions):
        """Extract solution and generate detailed reporting"""
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
        
        # --- REPORTING SECTION ---
        print("\n" + "="*40)
        print("          SCHEDULING REPORT")
        print("="*40)

        # 1. Workload Distribution
        print("\n[1] WORKLOAD DISTRIBUTION")
        sorted_counts = sorted(volunteer_counts.items(), key=lambda x: x[1], reverse=True)
        for name, count in sorted_counts:
            # Find volunteer info for context
            vol_info = next(v for v in self.volunteers if v['name'] == name)
            print(f" - {name:18} | Shifts: {count} | Team: {vol_info['team']:8} | Freq: {vol_info['frequency']}")

        # 2. Volunteers NOT Scheduled
        print("\n[2] VOLUNTEERS NOT SCHEDULED")
        not_scheduled = [v for v in self.volunteers if v['name'] not in scheduled_names]
        
        if not not_scheduled:
            print(" - Everyone was scheduled!")
        else:
            for v in not_scheduled:
                reason = ""
                if v['team'] == 'sub':
                    reason = "(Excluded: Team is 'Sub')"
                elif not v['priority'] and v['frequency'] == 'Monthly':
                    reason = "(Not needed for this rotation)"
                else:
                    reason = "(No available slots match constraints/priority)"
                
                print(f" - {v['name']:18} | Team: {v['team']:8} | {reason}")

        # 3. Unfilled Summary
        print("\n[3] COVERAGE SUMMARY")
        if total_unfilled == 0:
            print(" ✅ All positions filled successfully.")
        else:
            print(f" ⚠️  WARNING: {total_unfilled} positions were left UNFILLED.")

        print("="*40 + "\n")
        return schedule

def main():
    scheduler = VolunteerScheduler('volunteers.xlsx')
    scheduler.load_volunteers()
    scheduler.generate_schedule_dates(target_month=2, target_year=2026)
    res = scheduler.solve_schedule()
    if res: pd.DataFrame(res).to_excel('schedule_output.xlsx', index=False)

if __name__ == '__main__':
    main()