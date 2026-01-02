from typing import Any
import webcolors
import gspread
import cforces
import asyncio
import aiohttp
import string
import os


def find_intermediate_color(color1, color2, percentage):
    if not 0.0 <= percentage <= 1.0:
        raise ValueError("Percentage must be between 0.0 and 1.0")

    out = []
    for x in range(3):
        out.append(
            max(0, min(255, int(color1[x] + percentage * (color2[x] - color1[x]))))
        )

    return tuple(out)


colors = {"done": "#00FF00", "pending": "#FF0000"}


async def gather_data(contest_id: int, raw_problems: str, users: list[str]) -> Any:
    problems = []
    x = len(raw_problems) - 1
    while x >= 0:
        if raw_problems[x] in string.digits:
            x -= 1
            problems.append(raw_problems[x : x + 2])
        else:
            problems.append(raw_problems[x])
        x -= 1
    problems.reverse()

    # User: [Problem: [Non-]Solved], [Solve-Count]
    data = {}
    async with aiohttp.ClientSession() as sess:
        API = cforces.Client(sess)
        standings = await API.contest_standings(
            contest_id, handles=users, show_unofficial=True
        )

        std_info = []
        indexes = []

        for problem in standings.problems:
            if problem.index in problems:
                std_info.append(False)
            indexes.append(problem.index)

        if raw_problems == "AK":
            problems = indexes
            std_info = [False for _ in indexes]

        for row in standings.rows:
            if len(row.party.members) > 1:
                continue

            handle = row.party.members[0].handle.lower()
            if handle not in data:
                data[handle] = {"problems": std_info.copy()}

            for id, result in enumerate(row.problem_results):
                if not result.points:
                    continue
                if indexes[id] in problems:
                    data[handle]["problems"][id] = True

    for user in users:
        if user not in data:
            data[user] = {"problems": std_info}

    for handle in data.keys():
        solved = 0
        for status in data[handle]["problems"]:
            solved += int(status)
        data[handle]["solved"] = solved

    return data


def main():
    gc = gspread.service_account()
    wks = gc.open_by_key(os.environ["GS_SHEET_KEY"]).sheet1

    table_ref = None
    for x in range(1, 20):
        if wks.cell(x, 1).value == "Competencia":
            table_ref = x

    if not table_ref:
        raise ValueError("No table was found")

    contests_col = 4
    contests = []

    value = wks.cell(table_ref, contests_col, value_render_option="FORMULA").value
    if value:
        contests.append(value)

    while True:
        contests_col += 1
        value = wks.cell(table_ref, contests_col, value_render_option="FORMULA").value
        if not value:
            break
        contests.append(value)

    for idx in range(len(contests)):
        contest = contests[idx].replace(" ", "").split(",")
        contests[idx] = (contest[0][12:-1].rsplit("/", 1)[-1], contest[1][1:-2])

    users_row = table_ref + 1
    users = []
    while True:
        value = wks.cell(users_row, 1).value
        if not value:
            break
        users.append(value.lower())
        users_row += 1

    full_contests = {}
    for idx in range(len(contests)):
        contest_col = 4 + idx
        data = asyncio.run(gather_data(int(contests[idx][0]), contests[idx][1], users))

        cells = wks.range(
            table_ref + 1, contest_col, table_ref + 1 + len(users), contest_col
        )
        for id, user in enumerate(users):
            cells[id].value = (
                "=SPARKLINE("
                + str(data[user]["solved"])
                + ',{"charttype","bar";"max",'
                + str(len(data[user]["problems"]))
                + ';"color1","'
                + webcolors.rgb_to_hex(
                    find_intermediate_color(
                        webcolors.hex_to_rgb(colors["pending"]),
                        webcolors.hex_to_rgb(colors["done"]),
                        data[user]["solved"] / len(data[user]["problems"]),
                    )
                )
                + '"})'
            )

            if data[user]["solved"] == len(data[user]["problems"]):
                full_contests[user] = full_contests.get(user, 0) + 1

        wks.update_cells(
            cells, value_input_option=gspread.utils.ValueInputOption.user_entered
        )

    for id, user in enumerate(users):
        user_row = table_ref + id + 1
        wks.update_cell(user_row, 3, len(contests) - full_contests.get(user, 0) - 2)


if __name__ == "__main__":
    main()
