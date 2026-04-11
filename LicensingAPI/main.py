import os
import json
import pyodbc
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Optional

app = FastAPI(title="Licensing API")

# Dependency or utility to get DB connection
def get_db_connection():
    conn_str = os.getenv("DB_CONNECTION_STRING")
    print("DB_CONNECTION_STRING:", conn_str)
    if not conn_str:
        raise ValueError("DB_CONNECTION_STRING not set.")
    conn = pyodbc.connect(conn_str)

    conn.execute("SET NOCOUNT ON") 
    return conn

# --- Pydantic Models for Input/Output Validation ---

class AuthRequest(BaseModel):
    email: str
    password: str
    hardware_id: str

class DeviceInfo(BaseModel):
    device_id: int
    hardware_id: str
    added_at: str

class AuthResponse(BaseModel):
    status: str
    is_premium: Optional[bool] = None
    user_id: Optional[int] = None
    plan_name: Optional[str] = None
    allowed_modules: Optional[List[str]] = None
    registered_devices: Optional[List[DeviceInfo]] = None

class RemoveRequest(BaseModel):
    email: str
    password: str
    device_id: int

class RemoveResponse(BaseModel):
    status: str

class RegisterTrialRequest(BaseModel):
    full_name: str
    email: str
    mobile: str
    username: str
    password: str

class RegisterTrialResponse(BaseModel):
    status: str  # SUCCESS | ERROR_EMAIL_EXISTS | ERROR_MOBILE_EXISTS | ERROR_USERNAME_EXISTS

# --- Endpoints ---

@app.post("/authenticate", response_model=AuthResponse)
def authenticate_device(req: AuthRequest):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = "{CALL sp_AuthenticateWithDeviceManagement(?, ?, ?)}"
        cursor.execute(query, (req.email, req.password, req.hardware_id))

        if cursor.description is None:
            raise HTTPException(status_code=500, detail="No result set returned from SP")

        row = cursor.fetchone()
        if not row:
            return AuthResponse(status="ERROR_NO_RESPONSE")

        columns = [column[0] for column in cursor.description]
        result_dict = dict(zip(columns, row))
        resp_status = result_dict.get("Response")

        raw_modules = result_dict.get("AllowedModules")
        allowed_modules = None
        if raw_modules:
            try:
                allowed_modules = json.loads(raw_modules)
            except Exception:
                allowed_modules = None

        response = AuthResponse(
            status=resp_status,
            is_premium=result_dict.get("IsPremium"),
            user_id=result_dict.get("UserID"),
            plan_name=result_dict.get("PlanName"),
            allowed_modules=allowed_modules,
        )

        if resp_status == "LIMIT_REACHED" and cursor.nextset():
            if cursor.description:
                device_rows = cursor.fetchall()
                dev_cols = [column[0] for column in cursor.description]
                response.registered_devices = [
                    DeviceInfo(
                        device_id=d[0],
                        hardware_id=d[1],
                        added_at=str(d[2])
                    ) for d in device_rows
                ]

        conn.commit()
        return response
    except Exception as e:
        print(f"Error detail: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()


@app.post("/remove_device", response_model=RemoveResponse)
def remove_device(req: RemoveRequest):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = "{CALL sp_RemoveDevice(?, ?, ?)}"
        cursor.execute(query, (req.email, req.password, req.device_id))

        row = cursor.fetchone()
        if not row:
            return RemoveResponse(status="ERROR_NO_RESPONSE")

        conn.commit()
        return RemoveResponse(status=row[0])

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Database query failed: {str(e)}")
    finally:
        conn.close()


@app.post("/check_session", response_model=AuthResponse)
def check_session(req: AuthRequest):
    """
    Silently verifies that the device is still registered for the user.
    Used for auto-login on app startup. Returns SESSION_VALID or SESSION_INVALID.
    Never auto-registers a new device — admin remote logout is respected.
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = "{CALL sp_CheckDeviceSession(?, ?, ?)}"
        cursor.execute(query, (req.email, req.password, req.hardware_id))

        if cursor.description is None:
            raise HTTPException(status_code=500, detail="No result set returned from SP")

        row = cursor.fetchone()
        if not row:
            return AuthResponse(status="SESSION_INVALID")

        columns = [column[0] for column in cursor.description]
        result_dict = dict(zip(columns, row))
        resp_status = result_dict.get("Response")

        raw_modules = result_dict.get("AllowedModules")
        allowed_modules = None
        if raw_modules:
            try:
                allowed_modules = json.loads(raw_modules)
            except Exception:
                allowed_modules = None

        conn.commit()
        return AuthResponse(
            status=resp_status,
            is_premium=result_dict.get("IsPremium"),
            user_id=result_dict.get("UserID"),
            plan_name=result_dict.get("PlanName"),
            allowed_modules=allowed_modules,
        )
    except Exception as e:
        print(f"Error detail: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()


@app.post("/register_trial", response_model=RegisterTrialResponse)
def register_trial(req: RegisterTrialRequest):
    """
    Registers a new trial user. Checks email, mobile and username uniqueness
    via stored procedure, then inserts with a 7-day TrialExpiryDate.
    """
    try:
        conn = get_db_connection()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Database connection failed: {str(e)}")

    try:
        cursor = conn.cursor()
        query = "{CALL sp_RegisterTrialUser(?, ?, ?, ?, ?)}"
        params = (req.full_name, req.email, req.mobile, req.username, req.password)

        cursor.execute(query, params)
        # Fetch result only if result set exists
        if cursor.description is not None:
            row = cursor.fetchone()
        else:
            row = None

        if not row:
            return RegisterTrialResponse(status="ERROR_NO_RESPONSE")

        return RegisterTrialResponse(status=row[0])

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Database query failed: {str(e)}")
    finally:
        conn.close()


if __name__ == "__main__":
    import uvicorn
    # To run locally: uvicorn main:app --reload
    uvicorn.run(app, host="0.0.0.0", port=8000)
