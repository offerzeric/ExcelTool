from flask import Flask, json
from backend.code_in_424.code_in_424_controller import bp as code_in_424_bp
from werkzeug.exceptions import HTTPException
from flask_cors import CORS
#创建app入口
def create_app():
    app = Flask(__name__)   
    app.register_blueprint(code_in_424_bp)
    # app.register_blueprint(auth.bp)
    return app


app = create_app();
#解决跨域问题
CORS(app);


#全局的异常处理
@app.errorhandler(HTTPException)
def handle_http_exceptions(exception):
    response = exception.get_response()
    response.content_type = "application/json"
    response.status_code = 200;
    # replace the body with JSON
    response.data = json.dumps({
        "status": exception.code,
        "category": exception.name,
        "reason": exception.description,
        "flag": 0
    })
    return response;


if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0",port=9092)
