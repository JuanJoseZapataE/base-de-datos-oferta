import uvicorn

if __name__ == '__main__':
    uvicorn.run('fastapi_app.main:app', host='0.0.0.0', port=8000, reload=True)


# uvicorn main:app --reload
# .venv\Scripts\activate.bat