const fs = require('fs');
const vm = require('vm');
const assert = require('assert/strict');

const backend = {
  Utilities: {
    formatDate: (d, tz, fmt) => {
      const pad = n => String(n).padStart(2, '0');
      return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    },
    getUuid: () => Math.random().toString(36).slice(2, 10)
  },
  Session: {
    getScriptTimeZone: () => 'Asia/Bangkok'
  }
};
vm.createContext(backend);
vm.runInContext(fs.readFileSync('src/Code.js', 'utf8'), backend);

// Mock database and sheets
const sheets = {
  Employees: [
    { id: 'EMP-1', name: 'สมชาย ช่างทอง', position: 'ช่างเทคนิค', type: 'FULL_TIME', wageRate: 18000, otRate: 150, active: true },
    { id: 'EMP-2', name: 'วิชัย ปานสุข', position: 'โฟร์แมน', type: 'FULL_TIME', wageRate: 25000, otRate: 200, active: true }
  ],
  Attendance: [],
  Overtime: [],
  SalaryAdvances: [],
  SalaryPayments: [],
  Projects: [{ id: 'PRJ-1', name: 'โครงการ โกสุมพิสัย' }],
  AuditLogs: []
};

backend.db_ = () => ({});
backend.readSheet_ = (_, name) => sheets[name] || [];
backend.rows_ = (name) => sheets[name] || [];
backend.append_ = (name, row) => {
  const schema = vm.runInContext('APP.sheets', backend)[name];
  assert.ok(schema, 'Schema must exist for ' + name);
  const obj = {};
  schema.forEach((col, idx) => {
    obj[col] = row[idx] !== undefined ? row[idx] : '';
  });
  sheets[name].push(obj);
};
backend.updateRow_ = (name, id, patch) => {
  const idx = sheets[name].findIndex(x => x.id === id);
  assert.ok(idx !== -1, 'Row to update must exist: ' + id);
  sheets[name][idx] = { ...sheets[name][idx], ...patch };
};
backend.deleteRow_ = (name, id) => {
  const idx = sheets[name].findIndex(x => x.id === id);
  assert.ok(idx !== -1, 'Row to delete must exist: ' + id);
  sheets[name].splice(idx, 1);
};
backend.audit_ = () => {};

const user = { id: 'U-1', name: 'Admin', role: 'ADMIN', permissions: ['EMPLOYEE_MANAGE', 'OT_EDIT'] };

// Test 1: Check-in attendance
console.log('Test 1: Check-in attendance');
const checkInResult = backend.checkInAttendance_(user, {
  employeeId: 'EMP-1',
  date: '2026-10-04',
  checkInTime: '08:00',
  locationName: 'สำนักงานใหญ่ MB Solar',
  latitude: '16.123456',
  longitude: '102.654321',
  photoData: 'data:image/jpeg;base64,mockphotodata'
});

assert.ok(checkInResult.partial, 'Must return partial payload');
assert.equal(checkInResult.attendance.length, 1);
const att1 = checkInResult.attendance[0];
assert.equal(att1.employeeId, 'EMP-1');
assert.equal(att1.employeeName, 'สมชาย ช่างทอง');
assert.equal(att1.checkInTime, '08:00');
assert.equal(att1.locationName, 'สำนักงานใหญ่ MB Solar');
assert.equal(att1.latitude, '16.123456');
assert.equal(att1.longitude, '102.654321');
assert.equal(att1.status, 'CHECKED_IN');
console.log('PASS Test 1: Check-in successful');

// Test 2: Check-out attendance with OT
console.log('Test 2: Check-out attendance with OT');
const checkOutResult = backend.checkOutAttendance_(user, {
  id: att1.id,
  checkOutTime: '19:30',
  otHours: 2.5,
  note: 'เก็บงานระบบ inverter เรียบร้อย'
});

assert.equal(checkOutResult.attendance.length, 1);
const att1Updated = checkOutResult.attendance[0];
assert.equal(att1Updated.checkOutTime, '19:30');
assert.equal(att1Updated.otHours, 2.5);
// otRate for EMP-1 is 150 -> 2.5 * 150 = 375
assert.equal(att1Updated.otAmount, 375);
assert.equal(att1Updated.status, 'CHECKED_OUT');
console.log('PASS Test 2: Check-out with OT successful');

// Test 3: Save attendance with penalty (late / forgot check-in)
console.log('Test 3: Save attendance with penalty');
const penaltyResult = backend.saveAttendance_(user, {
  employeeId: 'EMP-2',
  date: '2026-10-04',
  checkInTime: '09:45',
  checkOutTime: '17:00',
  penaltyAmount: 100,
  penaltyReason: 'มาสาย 45 นาที',
  note: 'รถติด'
});

assert.equal(penaltyResult.attendance.length, 2);
const att2 = penaltyResult.attendance.find(x => x.employeeId === 'EMP-2');
assert.equal(att2.penaltyAmount, 100);
assert.equal(att2.penaltyReason, 'มาสาย 45 นาที');
assert.equal(att2.status, 'CHECKED_OUT');
console.log('PASS Test 3: Penalty recorded correctly');

// Test 4: Delete attendance
console.log('Test 4: Delete attendance');
const deleteResult = backend.deleteAttendance_(user, { id: att2.id });
assert.equal(deleteResult.attendance.length, 1);
assert.equal(deleteResult.attendance[0].id, att1.id);
console.log('PASS Test 4: Delete attendance successful');

// Test 5: Verify startupOnly returns [] for attendance
console.log('Test 5: StartupOnly bootstrap');
backend.monthInfo_ = () => ({ key: '2026-10', dueDate: '2026-10-31' });
backend.permissions_ = () => ['EMPLOYEE_MANAGE'];
backend.allowed_ = () => true;
backend.publicUser_ = u => u;

const startupBootstrap = backend.bootstrap_(user, true);
assert.equal(startupBootstrap.attendance.length, 0);
const fullBootstrap = backend.bootstrap_(user, false);
assert.equal(fullBootstrap.attendance.length, 1);
console.log('PASS Test 5: Startup optimization verified');

console.log('\nALL ATTENDANCE TESTS PASSED!');
