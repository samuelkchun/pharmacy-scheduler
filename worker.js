self.onmessage = function(e) {
  const { staff, perDiem, perDiemMax, personData, year, month, roles } = e.data;
  const localErrors = [];

  const numDays = new Date(year, month - 1, 0).getDate();
  const dates = [];
  const dow = [];
  for (let d = 1; d <= numDays; d++) {
    const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    dates.push(dateStr);
    const dateObj = new Date(year, month - 1, d);
    dow.push(dateObj.getDay()); // 0 Sun, 6 Sat
  }

  const required = dates.map((_, i) => {
    const isTueWed = dow[i] === 2 || dow[i] === 3; // Tue=2, Wed=3
    return { B: 5, S: isTueWed ? 3 : 2, L2: 1 };
  });

  const preAssign = {};
  const unavailable = {};
  const availablePerDiem = {};
  [...staff, ...perDiem].forEach(name => {
    unavailable[name] = [];
    if (perDiem.includes(name)) availablePerDiem[name] = [];
    dates.forEach(date => {
      let val = personData[name]?.[date] || '';
      if (val.startsWith('Prefer-')) {
        const role = val.replace('Prefer-', '');
        if (!preAssign[date]) preAssign[date] = [];
        preAssign[date].push({ name, role });
      }
      if (val === 'unavailable' || val === 'V' || val === 'P' || val === 'VAL') {
        unavailable[name].push(date);
      }
      if (perDiem.includes(name)) {
        if (val === 'available' || val.startsWith('Prefer-') || val === 'AM' || val === 'PM' || val === 'VAL') {
          availablePerDiem[name].push(date);
        }
      }
    });
  });

  let sched = {};
  dates.forEach(d => sched[d] = []);
  Object.keys(preAssign).forEach(date => {
    if (dates.includes(date)) {
      preAssign[date].forEach(as => {
        sched[date].push(as);
      });
    }
  });

  dates.forEach((date, i) => {
    sched[date].forEach(as => {
      if (roles.includes(as.role)) {
        required[i][as.role]--;
      }
    });
  });

  const unavailSet = {};
  staff.forEach(s => unavailSet[s] = new Set(unavailable[s]));
  const availPerSet = {};
  perDiem.forEach(p => availPerSet[p] = new Set(availablePerDiem[p]));

  for (let date in sched) {
    sched[date].forEach(as => {
      const isStaff = staff.includes(as.name);
      if (!isStaff && as.role !== 'B') {
        localErrors.push(`Invalid pre-assign: ${as.name} can't do ${as.role} on ${date}`);
      }
      if (isStaff && unavailSet[as.name].has(date)) {
        localErrors.push(`Invalid pre-assign: ${as.name} unavailable on ${date}`);
      }
      if (!isStaff && !availPerSet[as.name].has(date)) {
        localErrors.push(`Invalid pre-assign: ${as.name} not available on ${date}`);
      }
    });
  }

  let totalS = 0;
  required.forEach(r => totalS += r.S);

  const numStaff = staff.length;
  if (numStaff === 0) {
    localErrors.push('No staff added');
    self.postMessage({ schedule: sched, errors: localErrors });
    return;
  }
  const baseS = Math.floor(totalS / numStaff);
  const extraS = totalS % numStaff;
  const targetS = staff.map((_, i) => baseS + (i < extraS ? 1 : 0));

  const currentS = staff.map(() => 0);
  staff.forEach((s, si) => {
    dates.forEach(date => {
      if (sched[date].some(as => as.name === s && as.role === 'S')) {
        currentS[si]++;
      }
    });
  });

  staff.forEach((s, si) => {
    let remainingS = targetS[si] - currentS[si];
    while (remainingS > 0) {
      let blockSize = Math.min(3, remainingS);
      let found = false;
      for (let start = 0; start <= numDays - blockSize; start++) {
        let canAssign = true;
        for (let k = 0; k < blockSize; k++) {
          const day = start + k;
          const date = dates[day];
          if (unavailSet[s].has(date) || required[day].S <= 0 || sched[date].some(as => as.name === s)) {
            canAssign = false;
            break;
          }
        }
        if (canAssign) {
          for (let k = 0; k < blockSize; k++) {
            const day = start + k;
            const date = dates[day];
            sched[date].push({ name: s, role: 'S' });
            required[day].S--;
          }
          remainingS -= blockSize;
          found = true;
          break;
        }
      }
      if (!found && blockSize > 1) {
        blockSize--;
      } else if (!found) {
        localErrors.push(`Can't assign all S for ${s}`);
        break;
      }
    }
  });

  staff.sort((a, b) => currentS[staff.indexOf(a)] - currentS[staff.indexOf(b)]);
  for (let i = 0; i < numDays; i++) {
    while (required[i].S > 0) {
      let assigned = false;
      for (let s of staff) {
        const date = dates[i];
        if (!unavailSet[s].has(date) && !sched[date].some(as => as.name === s)) {
          sched[date].push({ name: s, role: 'S' });
          required[i].S--;
          assigned = true;
          break;
        }
      }
      if (!assigned) {
        localErrors.push(`Can't assign S on ${dates[i]} - not enough staff`);
        break;
      }
    }
  }

  for (let i = 0; i < numDays; i++) {
    while (required[i].L2 > 0) {
      let minShifts = Infinity;
      let bestS = null;
      const date = dates[i];
      staff.forEach(s => {
        if (!unavailSet[s].has(date) && !sched[date].some(as => as.name === s)) {
          const totalShifts = dates.reduce((count, d) => count + (sched[d].some(as => as.name === s) || personData[s]?.[d] === 'V' || personData[s]?.[d] === 'P' ? 1 : 0), 0);
          if (totalShifts < minShifts) {
            minShifts = totalShifts;
            bestS = s;
          }
        }
      });
      if (bestS) {
        sched[date].push({ name: bestS, role: 'L2' });
        required[i].L2--;
      } else {
        localErrors.push(`Can't assign L2 on ${date} - not enough staff`);
        break;
      }
    }
  }

  weeks.forEach(week => {
    const workDaysInWeek = staff.map(s => week.reduce((count, wd) => count + (sched[wd.date].some(as => as.name === s) || personData[s]?.[wd.date] === 'V' || personData[s]?.[wd.date] === 'P' ? 1 : 0), 0));
    let needs = staff.map((s, si) => 5 - workDaysInWeek[si]);
    while (true) {
      let assigned = false;
      staff.sort((a, b) => needs[staff.indexOf(b)] - needs[staff.indexOf(a)]);
      for (let s of staff) {
        const need = needs[staff.indexOf(s)];
        if (need <= 0) continue;
        for (let wd of week) {
          const date = wd.date;
          const ri = wd.index;
          if (required[ri].B > 0 && !unavailSet[s].has(date) && !sched[date].some(as => as.name === s)) {
            sched[date].push({ name: s, role: 'B' });
            required[ri].B--;
            needs[staff.indexOf(s)]--;
            assigned = true;
            break;
          }
        }
        if (assigned) break;
      }
      if (!assigned) break;
    }
  });

  dates.forEach((date, i) => {
    ['B', 'S', 'L2'].forEach(role => {
      if (required[i][role] > 0) {
        localErrors.push(`Missing ${required[i][role]} ${role} shifts on ${date}`);
      }
    });
  });

  staff.forEach(s => {
    let consec = 0;
    let maxConsec = 0;
    dates.forEach(date => {
      if (sched[date].some(as => as.name === s) || personData[s]?.[date] === 'V' || personData[s]?.[date] === 'P') {
        consec++;
        maxConsec = Math.max(maxConsec, consec);
      } else {
        consec = 0;
      }
    });
    if (maxConsec > 6) {
      console.log(`${s} has ${maxConsec} consecutive days (including V/P)`);
    }
  });

  const updatedPersonData = { ...personData };
  dates.forEach(date => {
    sched[date].forEach(as => {
      if (!updatedPersonData[as.name]) updatedPersonData[as.name] = {};
      if (!updatedPersonData[as.name][date]) {
        updatedPersonData[as.name][date] = as.role;
      }
    });
  });

  self.postMessage({ schedule: sched, errors: localErrors });
};