from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from sqlalchemy import text
from functools import wraps
import jwt
from urllib.parse import quote_plus
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')

# MySQL Database Configuration
DB_HOST = os.getenv('DB_HOST', 'localhost')
DB_PORT = os.getenv('DB_PORT', '3306')
DB_USER = os.getenv('DB_USER', 'root')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'root')
DB_NAME = os.getenv('DB_NAME', 'auth_system')

encoded_password = quote_plus(DB_PASSWORD)
app.config['SQLALCHEMY_DATABASE_URI'] = f'mysql+pymysql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_pre_ping': True,
    'pool_recycle': 300,
    'connect_args': {'charset': 'utf8mb4'}
}

# Email configuration
EMAIL_HOST = os.getenv('EMAIL_HOST', 'smtp.gmail.com')
EMAIL_PORT = int(os.getenv('EMAIL_PORT', '587'))
EMAIL_USER = os.getenv('EMAIL_USER', 'accllplogin@gmail.com')
EMAIL_PASS = os.getenv('EMAIL_PASS', 'accllp123')

db = SQLAlchemy(app)
CORS(app)

# Database Models
class User(db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    user_type = db.Column(db.Enum('admin', 'regular', name='user_types'), nullable=False)
    is_verified = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationship
    otps = db.relationship('OTP', backref='user', lazy=True, cascade='all, delete-orphan')
    
    def __repr__(self):
        return f'<User {self.email}>'

class OTP(db.Model):
    __tablename__ = 'otps'
    
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, index=True)
    otp_code = db.Column(db.String(6), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    expires_at = db.Column(db.DateTime, nullable=False)
    is_used = db.Column(db.Boolean, default=False, nullable=False)
    
    def __repr__(self):
        return f'<OTP {self.otp_code} for User {self.user_id}>'

# Utility Functions
def generate_otp():
    return str(random.randint(100000, 999999))

def send_email_otp(email, otp_code, name):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = email
        msg['Subject'] = "Your Login OTP"
        
        body = f"""
        Hi {name},
        
        Your OTP for login is: {otp_code}
        
        This OTP will expire in 10 minutes.
        
        Best regards,
        Your App Team
        """
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        text = msg.as_string()
        server.sendmail(EMAIL_USER, email, text)
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def token_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get('Authorization')
        if not token:
            return jsonify({'message': 'Token is missing'}), 401
        
        try:
            if token.startswith('Bearer '):
                token = token[7:]
            data = jwt.decode(token, app.config['SECRET_KEY'], algorithms=['HS256'])
            current_user = User.query.get(data['user_id'])
        except:
            return jsonify({'message': 'Token is invalid'}), 401
        
        return f(current_user, *args, **kwargs)
    return decorated

# API Routes
@app.route('/api/register', methods=['POST'])
def register():
    try:
        data = request.get_json()
        print(f"Received registration data: {data}")  # Add this for debugging
        
        # Validate required fields
        required_fields = ['email', 'password', 'full_name', 'phone', 'user_type']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'error': f'{field} is required'}), 400
        
        # Check if user already exists
        if User.query.filter_by(email=data['email']).first():
            return jsonify({'error': 'Email already registered'}), 400
        
        # Validate user_type (must be 'admin' or 'regular')
        if data['user_type'] not in ['admin', 'regular']:
            return jsonify({'error': 'user_type must be either "admin" or "regular"'}), 400
        
        try:
            # Create new user
            user = User(
                email=data['email'],
                password_hash=generate_password_hash(data['password']),
                full_name=data['full_name'],
                phone=data['phone'],
                user_type=data['user_type']
            )
            
            db.session.add(user)
            db.session.commit()
            print(f"User created successfully: {user.id}")  # Add this for debugging
            
            return jsonify({
                'message': 'User registered successfully',
                'user_id': user.id
            }), 201
            
        except Exception as db_error:
            db.session.rollback()  # Important: rollback on error
            print(f"Database error during registration: {db_error}")  # Detailed error logging
            return jsonify({'error': f'Database error: {str(db_error)}'}), 500
            
    except Exception as e:
        print(f"Unexpected error during registration: {str(e)}")  # Detailed error logging
        return jsonify({'error': str(e)}), 500
    

@app.route('/api/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        email = data.get('email')
        password = data.get('password')
        user_type = data.get('user_type')
        
        if not email or not password or not user_type:
            return jsonify({'error': 'Email, password, and user type are required'}), 400
        
        # Find user
        user = User.query.filter_by(email=email, user_type=user_type).first()
        
        if not user or not check_password_hash(user.password_hash, password):
            return jsonify({'error': 'Invalid credentials'}), 401
        
        # Generate and send OTP
        otp_code = generate_otp()
        expires_at = datetime.utcnow() + timedelta(minutes=10)
        
        # Remove old OTPs for this user
        OTP.query.filter_by(user_id=user.id, is_used=False).delete()
        
        # Create new OTP
        otp = OTP(
            user_id=user.id,
            otp_code=otp_code,
            expires_at=expires_at
        )
        
        db.session.add(otp)
        db.session.commit()
        
        # Send OTP via email (in production, you might want to use SMS)
        if send_email_otp(user.email, otp_code, user.full_name):
            return jsonify({
                'message': 'OTP sent successfully',
                'user_id': user.id,
                'requires_otp': True
            }), 200
        else:
            # For demo purposes, return the OTP in response
            return jsonify({
                'message': 'OTP generated (email failed)',
                'user_id': user.id,
                'requires_otp': True,
                'otp': otp_code  # Remove this in production
            }), 200
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/verify-otp', methods=['POST'])
def verify_otp():
    try:
        data = request.get_json()
        user_id = data.get('user_id')
        otp_code = data.get('otp_code')
        
        if not user_id or not otp_code:
            return jsonify({'error': 'User ID and OTP are required'}), 400
        
        # Find valid OTP
        otp = OTP.query.filter_by(
            user_id=user_id,
            otp_code=otp_code,
            is_used=False
        ).first()
        
        if not otp:
            return jsonify({'error': 'Invalid OTP'}), 401
        
        if datetime.utcnow() > otp.expires_at:
            return jsonify({'error': 'OTP has expired'}), 401
        
        # Mark OTP as used
        otp.is_used = True
        db.session.commit()
        
        # Get user details
        user = User.query.get(user_id)
        
        # Generate JWT token
        token = jwt.encode({
            'user_id': user.id,
            'email': user.email,
            'user_type': user.user_type,
            'exp': datetime.utcnow() + timedelta(hours=24)
        }, app.config['SECRET_KEY'], algorithm='HS256')
        
        return jsonify({
            'message': 'Login successful',
            'token': token,
            'user': {
                'id': user.id,
                'email': user.email,
                'full_name': user.full_name,
                'user_type': user.user_type
            }
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/profile', methods=['GET'])
@token_required
def get_profile(current_user):
    return jsonify({
        'user': {
            'id': current_user.id,
            'email': current_user.email,
            'full_name': current_user.full_name,
            'phone': current_user.phone,
            'user_type': current_user.user_type,
            'created_at': current_user.created_at.isoformat()
        }
    }), 200

@app.route('/api/resend-otp', methods=['POST'])
def resend_otp():
    try:
        data = request.get_json()
        user_id = data.get('user_id')
        
        if not user_id:
            return jsonify({'error': 'User ID is required'}), 400
        
        user = User.query.get(user_id)
        if not user:
            return jsonify({'error': 'User not found'}), 404
        
        # Generate new OTP
        otp_code = generate_otp()
        expires_at = datetime.utcnow() + timedelta(minutes=10)
        
        # Remove old OTPs
        OTP.query.filter_by(user_id=user.id, is_used=False).delete()
        
        # Create new OTP
        otp = OTP(
            user_id=user.id,
            otp_code=otp_code,
            expires_at=expires_at
        )
        
        db.session.add(otp)
        db.session.commit()
        
        # Send OTP
        if send_email_otp(user.email, otp_code, user.full_name):
            return jsonify({'message': 'OTP resent successfully'}), 200
        else:
            return jsonify({
                'message': 'OTP generated (email failed)',
                'otp': otp_code  # Remove this in production
            }), 200
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Initialize database
def create_tables():
    """Create database tables"""
    try:
        # Check if database exists first
        engine = db.create_engine(app.config['SQLALCHEMY_DATABASE_URI'])
        connection = engine.connect()
        connection.close()
        
        # Create tables
        db.create_all()
        print("Database tables created successfully!")
    except Exception as e:
        print(f"Error creating tables: {e}")
        print("Make sure your MySQL server is running and the database exists.")
        print(f"Database URI: {app.config['SQLALCHEMY_DATABASE_URI'].replace(DB_PASSWORD, '***')}")

@app.route('/api/test', methods=['GET'])
def test():
    return jsonify({
        'message': 'API is working',
        'timestamp': datetime.utcnow().isoformat()
    })

# 4. Add a separate database connection test
@app.route('/api/db-test', methods=['GET'])
def db_test():
    try:
        result = db.session.execute(text('SELECT 1 as test')).fetchone()
        return jsonify({
            'message': 'Database connection successful',
            'result': result[0] if result else None,
            'timestamp': datetime.utcnow().isoformat()
        })
    except Exception as e:
        return jsonify({
            'message': 'Database connection failed',
            'error': str(e),
            'timestamp': datetime.utcnow().isoformat()
        }), 500

if __name__ == '__main__':
    with app.app_context():
        create_tables()
        db.create_all()
    app.run(debug=True, host='0.0.0.0', port=5001)
