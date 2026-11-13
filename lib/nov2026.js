/* November 2026 drilling metres from the "Soosans drilling" WhatsApp shift reports (checked 13/11/2026).
   [date, shift, supervisor, block, [[rig, metres, hours from, hours to, problem, [[delay type, hours or null, what happened], …]], …]]
   Hours are only filled in where the report gives them. Shifts with no report in the chat are not listed. */
const B = "Breakdowns", S = "Service and maintenance", G = "Waiting for GET / drill bits", T = "Tramming", D = "Diesel / refuelling", W = "Meetings, transport, safety and weather", O = "Other";
const DRIFTER13 = (h = 12) => ["TDR 013", 0, "7979", "7979", "Drifter out for service (bucket yard)", [[B, h, "Drifter out for service – bucket yard"]]];
const ENGINE14 = ["TDR 014", 0, "", "", "Waiting for engine", [[B, 12, "Machine waiting for engine"]]];
const DUE15 = h => ["TDR 015", 0, "3424", "3424", "Drifter due for service (bucket yard)", [[B, h, "Drifter due for service – bucket yard"]]];
export const NOV_2026 = [
  ["2026-11-01", "Day", "Moses", "0323", [DRIFTER13(null),
    ["TDR 015", 85.7, "3413", "3421", "", [[S, 2, "Engineering servicing the machine"], [G, 0.5, "Waiting for GET"], [D, 0.5, "Refuelling diesel"]]]]],
  ["2026-11-01", "Night", "Katlego", "0323", [DRIFTER13(),
    ["TDR 015", 0, "3421", "3421", "No cut-off solenoid, E-stops not working (electrical)", [[W, 50 / 60, "Lightning alert"], [W, 0.5, "Transport to the pit"], [B, 6, "No cut-off solenoid, E-stops not working"]]], ENGINE14]],
  ["2026-11-02", "Day", "Moses", "0823", [DRIFTER13(null),
    ["TDR 015", 10.6, "3421", "3424", "", [[S, 6, "Weekly service"], [W, 1, "Lightning"], [T, 1, "Tramming from block 0323 to 0823"], [D, 0.5, "Refuelling diesel"]]]]],
  ["2026-11-02", "Night", "Katlego", "0823", [DRIFTER13(),
    ["TDR 015", 0, "3424", "3424", "No cut-off solenoid, E-stops not working (electrical)", [[W, 70 / 60, "Lightning alert"], [W, 0.5, "Transport to the pit"], [W, 0.75, "Risk assessment"], [B, 9.5, "No cut-off solenoid, E-stops not working"]]], ENGINE14]],
  ["2026-11-03", "Day", "Moses", "0823", [DRIFTER13(),
    ["TDR 015", 0, "3424", "3424", "Waiting for cut-off solenoid, E-stop not working", [[B, 12, "Waiting for cut-off solenoid, E-stop not working"]]]]],
  ["2026-11-03", "Night", "Katlego", "0823", [DRIFTER13(),
    ["TDR 015", 0, "3424", "3424", "No cut-off solenoid, E-stops not working (electrical)", [[B, 10, "No cut-off solenoid, E-stops not working"], [O, null, "No operator competent on the Soosan rigs on shift"]]], ENGINE14]],
  ["2026-11-04", "Day", "Moses", "0823", [DRIFTER13(),
    ["TDR 015", 10.7, "3424", "3425", "Solenoid and E-stop replaced, back on production 14:48", [[B, 8, "Engineering replacing solenoid and E-stop"]]]]],
  ["2026-11-04", "Night", "Yvonne", "0823", [DRIFTER13(),
    ["TDR 015", 98.4, "3425", "3433", "", [[O, 0.75, "Waiting for lighting plant"], [D, 40 / 60, "Out for diesel"]]]]],
  ["2026-11-05", "Day", "Moses", "0823", [DRIFTER13(),
    ["TDR 015", 0, "3433", "3434", "Drifter leaking oil – tramming to bucket yard to remove the drifter", [[B, 10, "Drifter leaking oil, under engineering"]]]]],
  ["2026-11-05", "Night", "Yvonne", "Bucket yard", [DRIFTER13(),
    ["TDR 015", 0, "3434", "3434", "Drifter leaking oil", [[B, 12, "Drifter leaking oil, under engineering"]]]]],
  ["2026-11-09", "Night", "Katlego", "Magazine / bucket yard", [DRIFTER13(), DUE15(12), ENGINE14]],
  ["2026-11-10", "Night", "Katlego", "Magazine / bucket yard", [DRIFTER13(), DUE15(12), ENGINE14]],
  ["2026-11-11", "Day", "Zandile", "Magazine / bucket yard", [DRIFTER13(), DUE15(12), ENGINE14]],
  ["2026-11-12", "Day", "Zandile", "Magazine / bucket yard", [DRIFTER13(), DUE15(12), ENGINE14]],
];
export const NOV_2026_TOTAL = 205.4;
export const NOV_2026_THROUGH = "12/11/2026 day shift";
