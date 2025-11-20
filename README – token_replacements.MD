# token_replacements.json – README  
**Onshape → WinCNC Post-Processor Rules**  
*ShopSabre / WinCNC Compatible G-code Converter*

---

### File Location
Place this file in the **same folder** Whoas `onshape-to-wincnc.pyw`  
(The script automatically detects and uses a local copy first — perfect for per-machine or per-job settings)

---

### Purpose
This JSON file **100% controls** what gets:
- Removed
- Replaced
- Commented out with `[ ]`

**No Python changes required** — just edit this file and you're done!

---

### Important: JSON Does NOT Allow Real Comments
Use keys that start with `_` — they are **ignored** by the script but perfect for documentation.

---

## File Structure & Example

```json
{
  "_INFO": "ShopSabre WinCNC post-processor rules – 2025 season",
  "_MACHINE": "ShopSabre PRO 408 – 10-station rack (positions 7–16)",

  "line_patterns": [
    {
      "match": "^[Oo]\\d+.*$",
      "action": "comment",
      "prefix": "[",
      "suffix": "]",
      "_note": "Program number lines → [O1000 (My Part)]"
    },
    {
      "match": "^%.*$",
      "action": "comment",
      "prefix": "[",
      "suffix": "]",
      "_note": "Old % delimiters → [%]"
    }
  ],

  "token_replacements": {
    "_remove_toolchange": "No ATC → strip M6",
    "M6": "",

    "_mist_control": "Mister on output port 1",
    "M7": "M11C1",
    "M9": "M12C1",

    "_flood_not_used": "Flood coolant disabled",
    "M8": "",

    "_standard_cleanup": "WinCNC doesn't understand these",
    "G17": "", "G40": "", "G80": "", "G49": "",

    "_tool_length_offset": "H-words ignored",
    "H\\d+": "",

    "_comment_line_numbers": "Keep visible for debugging",
    "(N\\d+)": "[\\1]",

    "_comment_tools_7_to_16": "Only rack tools are commented (preserved but inactive)",
    "(T(?:[7-9]|1[0-6]))": "[\\1]",

    "_leave_low_tools": "T1–T6 are manual/quick-change → leave active if needed"
  }
}

  ```
## Rule Reference Table
Rule,Effect,Before → After Example
"""M6"": """"",Remove tool change,T12 M6 → T12
"""M7"": ""M11C1"" / ""M9"": ""M12C1""",Convert mist coolant to WinCNC format,M7 → M11C1
"""G17"": """", ""G40"": """", ""G80"": """"",Remove unsupported modal codes,Removed
"""(N\\d+)"": ""[\\1]""",Comment line numbers,N150 → [N150]
"`""(T(?:[7-9]","1[0-6]))"": ""[\1]""`",Comment only tools 7–16 (your rack)
"""line_patterns"" → ""action"": ""comment""","Comment entire lines (O-lines, %)",O1000 → [O1000]

## Quick Customization Cheat Sheet
Goal,Change This Line
Comment all tool numbers,"""(T\\d+)"": ""[\\1]"""
Remove all tool numbers,"""T\\d+"": """""
Comment only T1–T6,"""(T[1-6])"": ""[\\1]"""
Use mister on port 3,"""M7"": ""M11C3"", ""M9"": ""M12C3"""
Keep M6 (you have an ATC),"Delete or comment out ""M6"": """""
Remove line numbers instead of commenting,"""(N\\d+)"": """""
Add flood coolant on port 2,"""M8"": ""M11C2"", add ""M9"": ""M12C2"" if needed"