import os
import urllib.parse
from sqlalchemy import create_engine, Column, Integer, String, Float, text
from sqlalchemy.orm import declarative_base, sessionmaker
from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

# Get the URL, default to the Supabase URL if Vercel doesn't load .env
raw_url = os.environ.get('POSTGRES_URL', 'postgresql://postgres.zrdufnladqvwedqkuyau:HEMkumar33#@aws-0-ap-northeast-1.pooler.supabase.com:6543/postgres')

# Force pg8000 driver for Vercel
if raw_url.startswith('postgresql://'):
    raw_url = raw_url.replace('postgresql://', 'postgresql://', 1)

# Fix password URL encoding if it contains special characters like #
if raw_url.startswith('postgresql://'):
    # Format is postgresql://user:password@host:port/db
    try:
        parts = raw_url.split('@')
        creds = parts[0].replace('postgresql://', '')
        user, password = creds.split(':', 1)
        safe_password = urllib.parse.quote_plus(urllib.parse.unquote_plus(password))
        DATABASE_URL = f"postgresql://{user}:{safe_password}@{parts[1]}"
    except Exception:
        DATABASE_URL = raw_url
else:
    DATABASE_URL = raw_url

engine = create_engine(DATABASE_URL)
Session = sessionmaker(bind=engine)
Base = declarative_base()

class StyleWeight(Base):
    __tablename__ = 'style_weights'
    id = Column(Integer, primary_key=True, autoincrement=True)
    style_name = Column(String, unique=True, nullable=False)
    weight = Column(Float, nullable=False)

class IdflStock(Base):
    __tablename__ = 'idfl_stock'
    id = Column(String, primary_key=True) # e.g. non_idfl_1
    tc_number = Column(String, nullable=False)
    products = Column(String, nullable=False)
    certified_weight = Column(Float, nullable=True)
    initial_weight = Column(Float, nullable=True)
    remaining_weight = Column(Float, nullable=False)
    sheet = Column(String, nullable=False)
    status = Column(String, nullable=False)

try:
    # Initialize database
    Base.metadata.create_all(engine)
    # Add column if missing
    with engine.connect() as conn:
        try: conn.execute(text("ALTER TABLE idfl_stock ADD COLUMN initial_weight FLOAT"))
        except: pass
        try: conn.execute(text("ALTER TABLE idfl_stock ADD COLUMN certified_weight FLOAT"))
        except: pass
        conn.commit()
except Exception as e:
    print(f"Warning: Could not initialize database tables: {e}")

def get_style_weights():
    with Session() as session:
        weights = session.query(StyleWeight).all()
        return {w.style_name: w.weight for w in weights}

def save_style_weights(weights_dict):
    with Session() as session:
        for style, weight in weights_dict.items():
            record = session.query(StyleWeight).filter_by(style_name=style).first()
            if record:
                record.weight = weight
            else:
                new_record = StyleWeight(style_name=style, weight=weight)
                session.add(new_record)
        session.commit()

def get_idfl_stock():
    with Session() as session:
        stock = session.query(IdflStock).all()
        result = []
        for s in stock:
            init_wt = s.initial_weight if s.initial_weight is not None else s.remaining_weight
            result.append({
                'id': s.id,
                'tc_number': s.tc_number,
                'products': s.products,
                'certified_weight': s.certified_weight if s.certified_weight is not None else init_wt,
                'initial_weight': init_wt,
                'used_weight': init_wt - s.remaining_weight,
                'remaining_weight': s.remaining_weight,
                'sheet': s.sheet,
                'status': s.status
            })
        return result

def save_idfl_stock(stock_list):
    with Session() as session:
        # First, delete all existing if we are doing a full overwrite 
        # (The UI handles full lists on init, and updates on consumption).
        # Actually, since app.py updates the whole list and calls save_idfl_stock(idfl_stock)
        # we can just upsert.
        
        # Clear existing table to ensure no stale data if list shrinks
        session.query(IdflStock).delete()
        
        for s in stock_list:
            new_record = IdflStock(
                id=s['id'],
                tc_number=s['tc_number'],
                products=s['products'],
                certified_weight=s.get('certified_weight', s.get('initial_weight', s['remaining_weight'])),
                initial_weight=s.get('initial_weight', s['remaining_weight']),
                remaining_weight=s['remaining_weight'],
                sheet=s['sheet'],
                status=s['status']
            )
            session.add(new_record)
        session.commit()
