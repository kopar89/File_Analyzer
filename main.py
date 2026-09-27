from pathlib import Path
import pandas as pd
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)


"""
DATABASE_URL = "postgresql://ps:123@localhost:5432/event_analytics"

engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)

Base = declarative_base()


class EventModel(Base):
    __tablename__ = "events"

    id = Column(
        Integer,
        primary_key=True,
        index=True
    )

    event = Column(
        String,
        nullable=False
    )

    time = Column(
        DateTime,
        nullable=False
    )


def get_db():
    db = SessionLocal()

    try:
        yield db
    finally:
        db.close()


class EventCreate(BaseModel):
    id: int
    event: str
    time: datetime
"""


@app.get("/")
def read_root():
    return FileResponse("main.html")


@app.post("/load_xl")
async def load(file: UploadFile = File(...)):

    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="Можно загружать только Excel-файлы"
        )

    file_path = UPLOAD_DIR / "students.xlsx"

    with open(file_path, "wb") as f:
        f.write(await file.read())

    return RedirectResponse(
        url="/show_btn",
        status_code=303
    )


@app.get("/show_btn")
def show():
    return FileResponse("metrik.html")


@app.get("/upload_db")
def texs():
    return FileResponse("tex.html")


@app.get("/group_by")
async def group_by():
    input_file = UPLOAD_DIR / "students.xlsx"
    output_file = UPLOAD_DIR / "grouped_students.xlsx"

    if not input_file.exists():
        raise HTTPException(
            status_code=404,
            detail="Excel-файл ещё не загружен"
        )

    df = pd.read_excel(input_file)

    required_columns = {
        "Группа",
        "ФИО",
        "Оценка",
        "Предмет"
    }

    missing_columns = required_columns - set(df.columns)

    if missing_columns:
        raise HTTPException(
            status_code=400,
            detail=(
                "В Excel отсутствуют столбцы: "
                + ", ".join(missing_columns)
            )
        )

    df = df[df["Оценка"].isin([5, "5"])]
    
    if df.empty:
        raise HTTPException(
            status_code=404,
            detail="В файле нет студентов с оценкой 5"
        )

    result = df[
        [
            "Группа",
            "ФИО",
            "Оценка",
            "Предмет"
        ]
    ]

    result = result.sort_values(
        by=[
            "Группа",
            "ФИО"
        ]
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        for group, group_df in result.groupby("Gруппа" if "Gруппа" in result.columns else "Группа"): # исправление для безопасности

            sheet_name = str(group)

            invalid_chars = [
                "\\",
                "/",
                "*",
                "?",
                ":",
                "[",
                "]"
            ]

            for char in invalid_chars:
                sheet_name = sheet_name.replace(char, "_")

            sheet_name = sheet_name[:31]

            group_df = group_df[
                [
                    "ФИО",
                    "Оценка",
                    "Предмет"
                ]
            ]

            group_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    return FileResponse(
        path=output_file,
        filename="grouped_students.xlsx",
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


@app.get("/avg")
async def avg():
    input_file = UPLOAD_DIR / "students.xlsx"
    output_file = UPLOAD_DIR / "average_students.xlsx"


    if not input_file.exists():
        raise HTTPException(
            status_code=404,
            detail="Excel-файл ещё не загружен"
        )

    df = pd.read_excel(input_file)

    #print(df.head())
    #print(df.columns)

    result = (
        df.groupby("Группа")["Оценка"].mean().reset_index()
    )

    result.to_excel(
        output_file,
        sheet_name="Лист1",
        index=False
    )

    return FileResponse(
        path=output_file,
        filename="average_students.xlsx"
    )


@app.get("/best_group")
async def best_group():
    input_file = UPLOAD_DIR / "students.xlsx"
    output_file = UPLOAD_DIR / "best_group.xlsx"


    if not input_file.exists():
        raise HTTPException(
            status_code=404,
            detail="Excel-файл ещё не загружен"
        )

    df = pd.read_excel(input_file)

    #print(df.head())
    #print(df.columns)

    result = (
        df.groupby("Группа")["Оценка"].mean().reset_index()
    )

    result = result.rename(columns={'Оценка': 'Средний балл'})
    result = result.sort_values(by=["Средний балл"], ascending=False)

    result.to_excel(
        output_file,
        sheet_name="Лист1",
        index=False
    )

    return FileResponse(
        path=output_file,
        filename="best_group.xlsx"
    )


@app.get("/group_count")
async def countr():

    input_file = UPLOAD_DIR / "students.xlsx"
    output_file = UPLOAD_DIR / "group_count.xlsx"

    if not input_file.exists():
        raise HTTPException(
            status_code=404,
            detail="Excel-файл ещё не загружен"
        )

    df = pd.read_excel(input_file)

    grade = df.groupby("Группа")["Оценка"].value_counts()
    result = grade.reset_index().rename(columns={'count': 'Количество'})


    grade_summary = df["Оценка"].value_counts().reset_index()
    grade_summary.columns = ["Оценка", "Количество"]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Лист1", index=False)
        grade_summary.to_excel(writer, sheet_name="Сводка", index=False)

    wb = load_workbook(output_file)
    ws = wb["Сводка"]

    pie = PieChart()
    pie.title = "Распределение оценок"


    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)   
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)   

    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    ws.add_chart(pie, "D2")  

    wb.save(output_file)

    return FileResponse(
        path=output_file,
        filename="group_count.xlsx"
    )

@app.get("/duty_student")
async def duty_check():
    input_file = UPLOAD_DIR / "students.xlsx"
    output_file = UPLOAD_DIR / "duty_student.xlsx"

    if not input_file.exists():
        raise HTTPException(
            status_code=404,
            detail="Excel-файл ещё не загружен"
        )

    df = pd.read_excel(input_file)

    result = df[df["Оценка"] == 2]


    result.to_excel(
            output_file,
            sheet_name="Лист1",
            index=False
        )
    
    return FileResponse(
            path=output_file,
            filename="duty_student.xlsx"
        )
    

"""
@app.post("/event")
def create_event(
    event_data: EventCreate,
    db: Session = Depends(get_db)
):

    db_event = EventModel(
        id=event_data.id,
        event=event_data.event,
        time=event_data.time
    )

    try:

        db.add(db_event)

        db.commit()

        db.refresh(db_event)

    except Exception as e:

        db.rollback()

        raise HTTPException(
            status_code=400,
            detail=f"Ошибка сохранения в БД: {str(e)}"
        )

    return {
        "status": "saved",
        "event": {
            "id": db_event.id,
            "event": db_event.event,
            "time": db_event.time
        }
    }
"""


"""
@app.get("/events")
def get_events(
    db: Session = Depends(get_db)
):

    result = db.execute(
        text("SELECT * FROM events")
    )

    events = [
        dict(row._mapping)
        for row in result
    ]

    return events
"""