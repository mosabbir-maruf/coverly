from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            form_data = {
                "HereNo": request.form["assignment_no"],
                "Here_Course_Code": request.form["course_code"],
                "Here_Course_Title": request.form["course_title"],
                "Here_TeacherName": request.form["teacher_name"],
                "designation": request.form["designation"],
                "Here_Teachers_Department_Name": request.form["teacher_dept"],
                "Here_StudentName": request.form["student_name"],
                "Here_StudentID": request.form["student_id"],
                "Here_Section": request.form["student_section"],
                "Here_DepartmentName": request.form["department_name"],
                "HereDate": request.form["submission_date"]
            }

            doc = Document("Assignment Cover Page.docx")
            for para in doc.paragraphs:
                for key, value in form_data.items():
                    if key in para.text:
                        for run in para.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, value)

            os.makedirs("output", exist_ok=True)
            date_str = datetime.now().strftime("%Y-%m-%d")
            docx_path = os.path.join("output", f"CoverPage_{date_str}.docx")
            doc.save(docx_path)

            return jsonify({"success": True, "file": docx_path})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)})

    return render_template("form.html")

@app.route("/download")
def download():
    from flask import request
    path = request.args.get("file")
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"✅ Server is running at http://0.0.0.0:{port}")
    app.run(debug=True, host="0.0.0.0", port=port)