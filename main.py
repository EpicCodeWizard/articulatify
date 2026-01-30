# startup script:
# python -m venv .venv
# .venv\Scripts\activate.bat
# pip install openpyxl requests gradio pytz
# python main.py

try:
    import openpyxl
except:
    import os
    os.system("pip install openpyxl requests gradio pytz")

import requests
import json
# from prompt_toolkit import prompt
# from prompt_toolkit.completion import WordCompleter
# from prompt_toolkit.validation import Validator, ValidationError

import openpyxl
from openpyxl.styles import Font, Alignment
from collections import defaultdict

import tempfile

import re
import os

import gradio as gr

# class OptionValidator(Validator):
#     def __init__(self, options):
#         self.options = options
#         super().__init__()

#     def validate(self, document):
#         text = document.text
#         if text not in self.options and not any(x for x in self.options if text.endswith(x)):
#             raise ValidationError(
#                 message="This input is not a valid option",
#                 cursor_position=len(text)
#             )


# result = prompt("Enter the source institution: ", completer=WordCompleter(names, ignore_case=True, match_middle=True), validator=OptionValidator(names))
# result = next(x for x in names if result.endswith(x))

# rawnames2 = requests.get("https://assist.org/api/institutions/" + str(sourceinst["id"]) + "/agreements").json()
# names2 = [x["institutionName"] for x in rawnames2]
# result2 = prompt("Enter the destination institution: ", completer=WordCompleter(names2, ignore_case=True, match_middle=True), validator=OptionValidator(names2))
# result2 = next(x for x in names2 if result2.endswith(x))

# destinst = next(x for x in rawnames2 if x["institutionName"] == result2)

# rawyears = requests.get("https://assist.org/api/AcademicYears").json()
# years = [str(x["FallYear"]) for x in rawyears if x["Id"] in destinst["receivingYearIds"]]
# year = prompt("Enter the year: ", completer=WordCompleter(years, ignore_case=True, match_middle=True), validator=OptionValidator(years), default=years[0])
# year = next(x for x in years if year.endswith(x))

# destyear = next(x for x in rawyears if str(x["FallYear"]) == year)


def flat_map(f, xs):
    ys = []
    for x in xs:
        ys.extend(f(x))
    return ys


def processarts(out):
    sir = {}

    for course in out:
        if not course["sendingArticulation"] or len(course["sendingArticulation"]["items"]) == 0:
            continue


        finval = ""
        if course["type"] == "Course":
            finval += "[" + course["course"]["prefix"] + " " + course["course"]["courseNumber"] + ": " + course["course"]["courseTitle"].strip() + "]"
        elif course["type"] == "Series":
            for art in course["series"]["courses"][:-1]:
                finval += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "] " + course["series"]["conjunction"].upper() + " "
            art = course["series"]["courses"][-1]
            finval += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "]"
        elif course["type"] == "Requirement":
            finval += "[REQUIREMENT: " + course["requirement"]["name"] + "]"
        elif course["type"] == "GeneralEducation":
            finval += "[GEN ED REQUIREMENT: " + course["generalEducationArea"]["name"] + "]"
        elif course["type"] == "Transferability":
            finval += "[TRANSFERABLE]"
        else:
            raise ValueError(f"Unknown course type: {course}")

        for sendart in course["sendingArticulation"]["items"][:-1]:
            finkey = ""
            if sendart["type"] == "Advisement": continue
            for art in sendart["items"][:-1]:
                finkey += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "] " + sendart["courseConjunction"].upper() + " "
            art = sendart["items"][-1]
            finkey += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "]"
            sir[finkey] = sir.get(finkey, []) + [finval]
        sendart = course["sendingArticulation"]["items"][-1]
        if sendart["type"] != "Advisement":
            finkey = ""
            for art in sendart["items"][:-1]:
                finkey += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "] " + sendart["courseConjunction"].upper() + " "
            art = sendart["items"][-1]
            finkey += "[" + art["prefix"] + " " + art["courseNumber"] + ": " + art["courseTitle"].strip() + "]"
            sir[finkey] = sir.get(finkey, []) + [finval]


    fin = {}

    for key, value in sir.items():
        lst = [x for x in list(set(value)) if "REQUIREMENT: " not in x] # otherwise can remove ts
        finlst = flat_map(lambda x: x.split(" OR "), lst)
        if len(finlst) > 0:
            for k in key.split(" OR "):
                toadd = fin.get(k, []) + finlst
                fin[k] = [s for s in toadd if not any(s in other and s != other for other in toadd)]

    return fin

def nameproc(name):
    return name.replace("University of California, ", "UC ").replace(" Los Angeles", "LA").replace("California State University, ", "CSU ")




def process_articulation(source_school, destination_schools, year_range, include_classname):

    logs = []

    def log(message):
        logs.append(message)
        return "\n".join(logs)

    yield None, log(f"[info] starting articulations scraping for {source_school}")

    destyear = next(x for x in rawyears if str(x["FallYear"]) == year_range.split("-")[0])
    myfin = {}
    classnames = {}

    yield None, log(f"[info] processing {len(destination_schools)} destination school(s) for {year_range} academic year")

    for result2 in destination_schools:
        yield None, log(f"[debug] fetching articulation data for {result2}")
        destinst = next(x for x in rawnames2 if x["institutionName"] == result2)

        lol = requests.get("https://assist.org/api/articulation/Agreements?Key=" + str(destyear["Id"]) + "/" + str(sourceinst["id"]) + "/to/" + str(destinst["institutionParentId"]) + "/AllMajors").json()

        if not lol.get("result"):
            yield None, log(f"[debug] trying reverse lookup for {result2}")
            lol = requests.get("https://assist.org/api/articulation/Agreements?Key=" + str(destyear["Id"]) + "/" + str(destinst["institutionParentId"]) + "/to/" + str(sourceinst["id"]) + "/AllMajors").json()
            if not lol.get("result"):
                yield None, log(f"[error] articulation not found for {result2}")
                continue
        out = [x["articulation"] for x in json.loads(lol["result"]["articulations"])]
        fin = processarts(out)

        torepl = {}
        for myc in fin:
            for myc2 in myc.split(" AND "):
                name = myc2[1:].split(": ")[0]
                torepl[" ".join(name.split(" ")[:-1])] = "".join(x for x in "".join(name.split(" ")[:-1]) if x.isalnum())

        myfin[result2] = []
        for course in fin:
            sirtoadd = " + ".join(sorted([x[1:].split(": ")[0] for x in course.split(" AND ")]))
            for cus in course.split(" AND "):
                templol = cus[1:].split(": ")[0]
                for k, v in torepl.items():
                    templol = templol.replace(k, v)
                classnames[templol] = cus[1:-1].split(": ")[1]
            for k, v in torepl.items():
                sirtoadd = sirtoadd.replace(k, v)
            myfin[result2].append(sirtoadd)
        myfin[result2] = list(sorted(list(set(myfin[result2])), key=lambda toproc: [toproc.split(" ")[0]] + [x for x in sum([[int(n), l] for part in toproc.replace("C100", "").split('+') for match in [re.findall(r'\d+|[A-Z]+', part.split()[-1])] for n, l in [(match[0], match[1] if len(match) > 1 else 'A')]], [])]))
        yield None, log(f"[info] completed articulations for {result2}, found {len(myfin[result2])} courses")

    yield None, log("[info] generating excel spreadsheet")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_empty = wb.create_sheet("\u200b", 0)
    ws_empty['A1'] = "Navigate to your desired department using the spreadsheet tabs at the bottom to view university course articulations."
    ws_empty['A1'].font = Font(size=18)
    ws_empty['A2'] = "Generated by Articulatify"
    ws_empty['A2'].font = Font(size=10)

    # Group courses by department
    dept_courses = defaultdict(set)
    for uc, courses in myfin.items():
        for course in courses:
            for sec in course.split(" + "):
                dept = sec.split(" ")[0]
                dept_courses[dept].add(course)

    ucs = list(myfin.keys())

    yield None, log(f"[info] creating {len(dept_courses)} department sheet(s)")

    for dept, courses in sorted(dept_courses.items()):
        ws = wb.create_sheet(dept)
        courses_sorted = list(sorted(list(set(courses)), key=lambda toproc: [toproc.split(" ")[0]] + [x for x in sum([[int(n), l] for part in toproc.replace("C100", "").split('+') for match in [re.findall(r'\d+|[A-Z]+', part.split()[-1])] for n, l in [(match[0], match[1] if len(match) > 1 else 'A')]], [])]))

        for col, course in enumerate(courses_sorted, 2):
            ws.cell(1, col, course)

        total_row = len(ucs) + 2
        for row, uc in enumerate(ucs, 2):
            ws.cell(row, 1, nameproc(uc))
            ws.cell(row+len(ucs)+3+(1 if include_classname else 0), 1, nameproc(uc))
            for col, course in enumerate(courses_sorted, 2):
                if course in myfin[uc]:
                    ws.cell(row, col, "✔")
                if include_classname: ws.cell(total_row+1, col, " + ".join([classnames[x] for x in course.split(" + ")]))

        ws.cell(total_row, 1, "Total")
        if include_classname: ws.cell(total_row+1, 1, "Course Name")
        ws.cell(total_row+len(ucs)+3+(1 if include_classname else 0), 1, "Total")
        if include_classname: ws.cell(total_row+len(ucs)+5, 1, "Course Name")
        for col in range(2, len(courses_sorted) + 2):
            ws.cell(total_row, col, f"=COUNTIF({openpyxl.utils.get_column_letter(col)}2:{openpyxl.utils.get_column_letter(col)}{len(ucs)+1}, \"✔\")")
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 14

        if include_classname: total_row += 1
        last_col = openpyxl.utils.get_column_letter(len(courses_sorted) + 1)
        ws.cell(total_row + 2, 2, f"=TRANSPOSE(SORT(TRANSPOSE({{B1:{last_col}1; B2:{last_col}{total_row}}}), ROWS(B1:{last_col}1)+MATCH(\"Total\", A2:A{total_row}, 0), FALSE))")

        ws.column_dimensions['A'].width = 20

    desired_filename = nameproc(source_school) + " Articulations.xlsx"

    yield None, log(f"[info] saving spreadsheet as {desired_filename}")

    # Save to temporary file
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, desired_filename)
    wb.save(tmp_path)

    yield tmp_path, log(f"[info] finished articulations scraping for {source_school}")



def get_source_schools():
    global rawnames
    rawnames = requests.get("https://assist.org/api/institutions").json()
    return sorted(list(set([x["names"][0]["name"] for x in rawnames])))

def get_destination_schools(source_school):
    global rawnames2, sourceinst
    sourceinst = next(x for x in rawnames if x["names"][0]["name"] == source_school)
    rawnames2 = requests.get("https://assist.org/api/institutions/" + str(sourceinst["id"]) + "/agreements").json()
    names2 = [x["institutionName"] for x in rawnames2]
    return gr.Dropdown(choices=sorted(list(set(names2))), multiselect=True, label="Destination Institutions")


def get_years(source_school, destination_schools):
    global rawyears
    lolfin = []
    for scol in destination_schools:
        destinst = next(x for x in rawnames2 if x["institutionName"] == scol)
        rawyears = requests.get("https://assist.org/api/AcademicYears").json()
        lolfin.append([str(x["FallYear"]) + "-" + str(x["FallYear"]+1) for x in rawyears if x["Id"] in destinst["receivingYearIds"]+destinst["sendingYearIds"]])
    return gr.Dropdown(choices=sorted(set.intersection(*map(set, lolfin)), reverse=True), label="Year")


with gr.Blocks(title="Articulatify") as iface:
    gr.Markdown("# Articulatify")
    gr.Markdown("### Generate bulk course articulations spreadsheets using Assist.org.")

    source = gr.Dropdown(choices=get_source_schools(), label="Source Institution", value=None)
    destination = gr.Dropdown(label="Destination Institutions", multiselect=True, visible=False)
    year = gr.Dropdown(label="Year", visible=False)
    include_classname = gr.Checkbox(label="Include Class Names", value=False)

    with gr.Row():
        submit = gr.Button("Generate")

    output = gr.File(label="Download Spreadsheet")

    logs = gr.Textbox(
        label="Logs",
        lines=10,
        max_lines=20,
        interactive=False,
        autoscroll=False,
        elem_id="logs-box"
    )

    source.change(
        fn=lambda: gr.Dropdown(visible=True),
        outputs=destination
    ).then(
        fn=get_destination_schools,
        inputs=source,
        outputs=destination
    )

    destination.change(
        fn=lambda: gr.Dropdown(visible=True),
        outputs=year
    ).then(
        fn=get_years,
        inputs=[source, destination],
        outputs=year
    )

    submit.click(fn=process_articulation, inputs=[source, destination, year, include_classname], outputs=[output, logs])


if __name__ == "__main__":
    iface.launch(pwa=True, css="""
    footer{display:none !important}
    #logs-box .generating {border:none !important}
    #logs-box textarea {scroll-behavior:auto !important}
    #logs-box.generating {border-color:var(--border-color-primary) !important}
    .generating {border-color:var(--border-color-primary) !important}
""", favicon_path="./lolicon.png", server_name="0.0.0.0")

# include course name
