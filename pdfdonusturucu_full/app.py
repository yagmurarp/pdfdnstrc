import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from converters.pdf_ops import pdf_to_docx
from converters.word_excel import docx_to_xlsx, xlsx_to_docx
ALLOWED_WORD={'.docx'}; ALLOWED_EXCEL={'.xlsx'}; ALLOWED_PDF={'.pdf'}
app=Flask(__name__); app.secret_key=os.environ.get('SECRET_KEY','dev'); app.config['UPLOAD_FOLDER']=os.path.abspath('uploads'); app.config['OUTPUT_FOLDER']=os.path.abspath('outputs'); app.config['MAX_CONTENT_LENGTH']=100*1024*1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True); os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def allowed(fn,exts): return os.path.splitext(fn)[1].lower() in exts

@app.route('/',methods=['GET'])
def index(): return render_template('index.html')

@app.route('/convert',methods=['POST'])
def convert():
    action=request.form.get('action'); f=request.files.get('file')
    if not action or not f or f.filename.strip()=='' :
        flash('Lütfen bir dosya seç ve işlem türünü gönder.'); return redirect(url_for('index'))
    filename=secure_filename(f.filename); src=os.path.join(app.config['UPLOAD_FOLDER'],filename); f.save(src)
    name,ext=os.path.splitext(filename); ext=ext.lower()
    try:
        if action=='pdf_to_word':
            if not allowed(filename,ALLOWED_PDF): raise ValueError('Sadece PDF kabul edilir.')
            out=os.path.join(app.config['OUTPUT_FOLDER'],f"{name}.docx"); pdf_to_docx(src,out); return send_file(out,as_attachment=True,download_name=f"{name}.docx")
        elif action=='word_to_pdf':
            flash('Word → PDF çevrimiçi sürümde kapalı (LibreOffice gerek). Masaüstünde çalışır.'); return redirect(url_for('index'))
        elif action=='pdf_to_excel':
            flash('PDF → Excel çevrimiçi sürümde kapalı (Ghostscript gerek). Masaüstünde çalışır.'); return redirect(url_for('index'))
        elif action=='excel_to_pdf':
            flash('Excel → PDF çevrimiçi sürümde kapalı (LibreOffice gerek). Masaüstünde çalışır.'); return redirect(url_for('index'))
        elif action=='word_to_excel':
            if not allowed(filename,ALLOWED_WORD): raise ValueError('Sadece DOCX kabul edilir.')
            out=os.path.join(app.config['OUTPUT_FOLDER'],f"{name}.xlsx"); docx_to_xlsx(src,out); return send_file(out,as_attachment=True,download_name=f"{name}.xlsx")
        elif action=='excel_to_word':
            if not allowed(filename,ALLOWED_EXCEL): raise ValueError('Sadece XLSX kabul edilir.')
            out=os.path.join(app.config['OUTPUT_FOLDER'],f"{name}.docx"); xlsx_to_docx(src,out); return send_file(out,as_attachment=True,download_name=f"{name}.docx")
        else:
            raise ValueError('Bilinmeyen işlem.')
    except Exception as e:
        app.logger.exception('Dönüşüm hatası'); flash(f'Hata: {e}'); return redirect(url_for('index'))
    finally:
        try: os.remove(src)
        except Exception: pass

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

