const ATTENDANCE_DATA_KEY = "schoolAttendanceData";
const DAILY_ATTENDANCE_KEY = "schoolDailyAttendance";

function saveSharedAttendanceData(data) {
  localStorage.setItem(ATTENDANCE_DATA_KEY, JSON.stringify(data));
}

function loadSharedAttendanceData() {
  try {
    return JSON.parse(localStorage.getItem(ATTENDANCE_DATA_KEY) || "null");
  } catch (error) {
    return null;
  }
}

function clearSharedAttendanceData() {
  localStorage.removeItem(ATTENDANCE_DATA_KEY);
}

function saveDailyAttendance(section, date, record) {
  const attendance = loadDailyAttendance();
  attendance[section] = attendance[section] || {};
  attendance[section][date] = record;
  localStorage.setItem(DAILY_ATTENDANCE_KEY, JSON.stringify(attendance));
}

function loadDailyAttendance() {
  try {
    return JSON.parse(localStorage.getItem(DAILY_ATTENDANCE_KEY) || "{}");
  } catch (error) {
    return {};
  }
}

function clearDailyAttendance() {
  localStorage.removeItem(DAILY_ATTENDANCE_KEY);
}
function findSharedSection(data, className, sectionName) {
  const keys = Object.keys(data || {});
  const classValue = String(className || "").trim().toLowerCase();
  const sectionValue = String(sectionName || "").trim().toLowerCase();
  const combined = `${classValue}-${sectionValue}`.replace(/\s+/g, "");

  return keys.find(key => key.toLowerCase().replace(/\s+/g, "") === combined)
    || keys.find(key => key.toLowerCase() === classValue)
    || keys[0]
    || "";
}

function setSharedSectionInputs(key, classInputId, sectionInputId) {
  const separatorIndex = key.lastIndexOf("-");
  if (separatorIndex < 0) return;
  document.getElementById(classInputId).value = key.slice(0, separatorIndex);
  document.getElementById(sectionInputId).value = key.slice(separatorIndex + 1);
}

function populateSharedSelectors(data, classSelect, sectionSelect) {
  const keys = Object.keys(data || {});
  const fallbackClasses = ["5th", "6th", "7th", "8th", "9th", "10th"];
  const classes = keys.length
    ? [...new Set(keys.map(key => key.slice(0, key.lastIndexOf("-"))))]
    : fallbackClasses;
  const currentClass = classes.includes(classSelect.value) ? classSelect.value : classes[0];

  classSelect.innerHTML = classes.map(value => `<option value="${value}">${value}</option>`).join("");
  classSelect.value = currentClass;

  const sections = keys.length
    ? keys.filter(key => key.slice(0, key.lastIndexOf("-")) === currentClass)
      .map(key => key.slice(key.lastIndexOf("-") + 1))
    : ["A", "B", "C", "D", "E", "F"];
  const currentSection = sections.includes(sectionSelect.value) ? sectionSelect.value : sections[0];

  sectionSelect.innerHTML = sections.map(value => `<option value="${value}">${value}</option>`).join("");
  sectionSelect.value = currentSection;
}