# Excel AI Chat Backend Architecture Documentation

## Table of Contents
1. [Executive Summary](#executive-summary)
2. [System Architecture Overview](#system-architecture-overview)
3. [Frontend Integration Specifications](#frontend-integration-specifications)
4. [Backend Application Structure](#backend-application-structure)
5. [Authentication & Security Framework](#authentication--security-framework)
6. [AI Integration Architecture](#ai-integration-architecture)
7. [Data Processing Pipeline](#data-processing-pipeline)
8. [API Design & Contracts](#api-design--contracts)
9. [Infrastructure & Deployment](#infrastructure--deployment)
10. [Development Environment Setup](#development-environment-setup)
11. [Testing Strategy](#testing-strategy)
12. [Monitoring & Observability](#monitoring--observability)
13. [Security Considerations](#security-considerations)
14. [Performance Optimization](#performance-optimization)
15. [Future Enhancements](#future-enhancements)

---

## Executive Summary

This document provides a comprehensive architectural blueprint for developing a Python backend service that integrates with the existing Excel Office.js task pane add-in. The backend system is designed to process Excel data context, provide AI-powered responses through OpenAI integration, and maintain enterprise-grade security and scalability.

### Key Design Principles
- **Microservices Architecture**: Modular, scalable design using FastAPI
- **Domain-Driven Design**: Clear separation of concerns across business domains
- **Security-First**: Office.js SSO integration with JWT validation
- **AI-Powered**: OpenAI integration with Excel context awareness
- **Production-Ready**: Comprehensive logging, monitoring, and deployment strategies

### Technology Stack
- **Framework**: FastAPI (Python 3.11+)
- **AI Integration**: OpenAI GPT-4/3.5-turbo
- **Data Processing**: Pandas, NumPy
- **Authentication**: Microsoft Identity Platform (MSAL)
- **Database**: PostgreSQL with SQLAlchemy ORM
- **Caching**: Redis
- **Deployment**: Docker, Azure Container Apps / Railway
- **Monitoring**: Prometheus, Grafana, Application Insights

---

## System Architecture Overview

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Excel Office.js Frontend                     │
│  ┌─────────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │   Task Pane     │  │   Chat Interface │  │  Excel Context  │ │
│  │   (HTML/CSS/JS) │  │   (Messages)     │  │  (Selected Data)│ │
│  └─────────────────┘  └──────────────────┘  └─────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
                                │
                                │ HTTPS/REST API
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Python Backend Services                    │
│                                                                 │
│  ┌─────────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │   API Gateway   │  │   Auth Service   │  │   AI Service    │ │
│  │   (FastAPI)     │  │   (JWT/SSO)      │  │   (OpenAI)      │ │
│  └─────────────────┘  └──────────────────┘  └─────────────────┘ │
│                                                                 │
│  ┌─────────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │  Excel Service  │  │   Chat Service   │  │  Cache Layer    │ │
│  │  (Data Proc.)   │  │   (Messages)     │  │   (Redis)       │ │
│  └─────────────────┘  └──────────────────┘  └─────────────────┘ │
│                                                                 │
│  ┌─────────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │   Database      │  │   Monitoring     │  │   Logging       │ │
│  │  (PostgreSQL)   │  │  (Prometheus)    │  │  (Structured)   │ │
│  └─────────────────┘  └──────────────────┘  └─────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
                                │
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                     External Services                           │
│  ┌─────────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │   OpenAI API    │  │  Microsoft Graph │  │   Azure AD      │ │
│  │   (GPT-4)       │  │   (Optional)     │  │   (Identity)    │ │
│  └─────────────────┘  └──────────────────┘  └─────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
```

### 3-Tier Architecture Implementation

#### 1. Presentation Layer (API Gateway)
- **FastAPI Routers**: Handle HTTP requests/responses
- **Request Validation**: Pydantic models for data validation
- **CORS Configuration**: Enable GitHub Pages frontend access
- **Rate Limiting**: Protect against abuse

#### 2. Business Logic Layer (Services)
- **Chat Service**: Message processing and conversation management
- **Excel Service**: Data parsing and analysis
- **AI Service**: OpenAI integration and prompt engineering
- **Auth Service**: Authentication and authorization

#### 3. Data Access Layer (Repositories)
- **Database Models**: SQLAlchemy ORM entities
- **Repository Pattern**: Data access abstraction
- **Cache Management**: Redis for performance optimization

---

## Frontend Integration Specifications

### Current Frontend Implementation

The existing Excel Office.js frontend is deployed at `https://mavencat.github.io/frontend/taskpane.html` and provides the following integration points:

#### API Call Structure
```javascript
// Current frontend API call (taskpane.js:182)
const API_URL = 'http://localhost:8000/api/chat';

const payload = {
    message: userMessage,
    context: excelContext,
    timestamp: new Date().toISOString()
};

const response = await fetch(API_URL, {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload)
});
```

#### Excel Context Data Structure
```javascript
// Excel context extracted from Office.js API
const excelContext = {
    selectedRange: {
        address: "A1:B10",
        values: [["Header1", "Header2"], ["Value1", "Value2"]],
        rowCount: 10,
        columnCount: 2
    },
    worksheet: {
        name: "Sheet1"
    }
};
```

#### Expected Response Format
```javascript
// Frontend expects this response structure
{
    "response": "AI's answer to the user's question"
}
```

### Frontend Security Model

#### CORS Requirements
```python
# Required CORS configuration for backend
origins = [
    "https://mavencat.github.io",  # Production frontend
    "https://localhost:3000",      # Development frontend
    "https://excel-wd.officeapps.live.com",  # Excel Online domains
    "https://office.live.com",     # Office 365 domains
]
```

#### Office.js Security Context
- Frontend runs in Excel task pane iframe
- Requires HTTPS for all external API calls
- Office.js provides SSO tokens for authenticated users
- Frontend has ReadWriteDocument permissions on Excel workbook

---

## Backend Application Structure

### Recommended Project Structure

```
excel-ai-backend/
├── app/
│   ├── __init__.py
│   ├── main.py                    # FastAPI application entry point
│   ├── config.py                  # Application configuration
│   ├── dependencies.py            # Shared dependencies
│   │
│   ├── api/                       # API layer
│   │   ├── __init__.py
│   │   ├── v1/
│   │   │   ├── __init__.py
│   │   │   ├── endpoints/
│   │   │   │   ├── __init__.py
│   │   │   │   ├── chat.py        # Chat endpoints
│   │   │   │   ├── auth.py        # Authentication endpoints
│   │   │   │   └── health.py      # Health check endpoints
│   │   │   └── api.py             # API router aggregation
│   │   └── middleware/
│   │       ├── __init__.py
│   │       ├── cors.py            # CORS middleware
│   │       ├── auth.py            # Authentication middleware
│   │       └── logging.py         # Request logging
│   │
│   ├── core/                      # Core business logic
│   │   ├── __init__.py
│   │   ├── auth/
│   │   │   ├── __init__.py
│   │   │   ├── jwt_handler.py     # JWT token validation
│   │   │   ├── office_sso.py      # Office.js SSO integration
│   │   │   └── microsoft_auth.py  # Microsoft Identity Platform
│   │   ├── ai/
│   │   │   ├── __init__.py
│   │   │   ├── openai_client.py   # OpenAI API client
│   │   │   ├── prompt_templates.py # AI prompt management
│   │   │   └── response_processor.py # AI response processing
│   │   ├── excel/
│   │   │   ├── __init__.py
│   │   │   ├── data_processor.py  # Excel data analysis
│   │   │   ├── context_parser.py  # Parse Excel context
│   │   │   └── insights_generator.py # Generate insights
│   │   └── chat/
│   │       ├── __init__.py
│   │       ├── message_handler.py # Message processing
│   │       ├── conversation.py    # Conversation management
│   │       └── session_manager.py # User session handling
│   │
│   ├── models/                    # Database models
│   │   ├── __init__.py
│   │   ├── base.py               # Base model classes
│   │   ├── user.py               # User models
│   │   ├── chat.py               # Chat/message models
│   │   └── session.py            # Session models
│   │
│   ├── schemas/                   # Pydantic schemas
│   │   ├── __init__.py
│   │   ├── chat.py               # Chat request/response schemas
│   │   ├── auth.py               # Authentication schemas
│   │   ├── excel.py              # Excel data schemas
│   │   └── common.py             # Common schemas
│   │
│   ├── services/                  # Business services
│   │   ├── __init__.py
│   │   ├── chat_service.py       # Chat business logic
│   │   ├── ai_service.py         # AI integration service
│   │   ├── excel_service.py      # Excel processing service
│   │   ├── auth_service.py       # Authentication service
│   │   └── cache_service.py      # Caching service
│   │
│   ├── repositories/              # Data access layer
│   │   ├── __init__.py
│   │   ├── base.py               # Base repository
│   │   ├── chat_repository.py    # Chat data access
│   │   ├── user_repository.py    # User data access
│   │   └── session_repository.py # Session data access
│   │
│   └── utils/                     # Utility functions
│       ├── __init__.py
│       ├── logging.py            # Logging utilities
│       ├── security.py           # Security utilities
│       ├── validation.py         # Data validation
│       └── exceptions.py         # Custom exceptions
│
├── tests/                         # Test suite
│   ├── __init__.py
│   ├── conftest.py               # Test configuration
│   ├── unit/                     # Unit tests
│   ├── integration/              # Integration tests
│   └── e2e/                      # End-to-end tests
│
├── deployment/                    # Deployment configurations
│   ├── docker/
│   │   ├── Dockerfile
│   │   └── docker-compose.yml
│   ├── kubernetes/
│   │   ├── deployment.yaml
│   │   └── service.yaml
│   └── azure/
│       └── container-apps.yaml
│
├── scripts/                       # Utility scripts
│   ├── start_dev.py              # Development server
│   ├── migrate_db.py             # Database migrations
│   └── seed_data.py              # Test data seeding
│
├── requirements/                  # Python dependencies
│   ├── base.txt                  # Base requirements
│   ├── dev.txt                   # Development requirements
│   └── prod.txt                  # Production requirements
│
├── .env.example                   # Environment variables template
├── .gitignore
├── README.md
├── pyproject.toml                 # Python project configuration
└── alembic.ini                   # Database migration configuration
```

### Core Application Files

#### main.py - Application Entry Point
```python
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api.v1.api import api_router
from app.core.config import settings
from app.middleware.cors import setup_cors
from app.middleware.logging import setup_logging

app = FastAPI(
    title="Excel AI Chat Backend",
    description="AI-powered chat backend for Excel Office.js add-in",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Setup middleware
setup_cors(app)
setup_logging(app)

# Include routers
app.include_router(api_router, prefix="/api/v1")

@app.get("/")
async def root():
    return {"message": "Excel AI Chat Backend API"}

@app.get("/health")
async def health_check():
    return {"status": "healthy", "version": "1.0.0"}
```

#### config.py - Configuration Management
```python
from pydantic_settings import BaseSettings
from typing import List, Optional

class Settings(BaseSettings):
    # Application
    APP_NAME: str = "Excel AI Chat Backend"
    DEBUG: bool = False
    API_V1_STR: str = "/api/v1"
    
    # Security
    SECRET_KEY: str
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    
    # CORS
    ALLOWED_ORIGINS: List[str] = [
        "https://mavencat.github.io",
        "https://localhost:3000"
    ]
    
    # Database
    DATABASE_URL: str
    
    # Redis
    REDIS_URL: str = "redis://localhost:6379"
    
    # OpenAI
    OPENAI_API_KEY: str
    OPENAI_MODEL: str = "gpt-4"
    OPENAI_MAX_TOKENS: int = 1000
    
    # Microsoft Identity
    AZURE_TENANT_ID: str
    AZURE_CLIENT_ID: str
    AZURE_CLIENT_SECRET: str
    
    # Logging
    LOG_LEVEL: str = "INFO"
    
    class Config:
        env_file = ".env"

settings = Settings()
```

---

## Authentication & Security Framework

### Office.js SSO Integration

#### JWT Token Validation Flow

```mermaid
sequenceDiagram
    participant Frontend as Excel Frontend
    participant Backend as Python Backend
    participant Azure as Azure AD
    participant OpenAI as OpenAI API

    Frontend->>Frontend: User sends message
    Frontend->>Azure: getAccessToken() via Office.js
    Azure-->>Frontend: JWT Token
    Frontend->>Backend: POST /api/chat with JWT
    Backend->>Backend: Validate JWT token
    Backend->>Backend: Extract user identity
    Backend->>OpenAI: Generate AI response
    OpenAI-->>Backend: AI response
    Backend-->>Frontend: Formatted response
```

#### JWT Token Validation Implementation

```python
# app/core/auth/jwt_handler.py
import jwt
from jwt import PyJWTError
from fastapi import HTTPException, status
from typing import Optional
import httpx
from app.core.config import settings

class JWTHandler:
    def __init__(self):
        self.azure_tenant_id = settings.AZURE_TENANT_ID
        self.azure_client_id = settings.AZURE_CLIENT_ID
        
    async def validate_office_token(self, token: str) -> dict:
        """Validate Office.js SSO token"""
        try:
            # Decode without verification first to get header
            unverified_header = jwt.get_unverified_header(token)
            kid = unverified_header.get('kid')
            
            # Get Microsoft's public keys
            jwks = await self._get_microsoft_jwks()
            
            # Find the correct key
            public_key = self._find_key(jwks, kid)
            
            # Verify and decode token
            payload = jwt.decode(
                token,
                public_key,
                algorithms=['RS256'],
                audience=self.azure_client_id,
                issuer=f"https://login.microsoftonline.com/{self.azure_tenant_id}/v2.0"
            )
            
            return payload
            
        except PyJWTError as e:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail=f"Token validation failed: {str(e)}"
            )
    
    async def _get_microsoft_jwks(self) -> dict:
        """Fetch Microsoft's JWKS endpoint"""
        jwks_url = f"https://login.microsoftonline.com/{self.azure_tenant_id}/discovery/v2.0/keys"
        
        async with httpx.AsyncClient() as client:
            response = await client.get(jwks_url)
            response.raise_for_status()
            return response.json()
    
    def _find_key(self, jwks: dict, kid: str) -> str:
        """Find public key from JWKS by key ID"""
        for key in jwks.get('keys', []):
            if key.get('kid') == kid:
                return jwt.algorithms.RSAAlgorithm.from_jwk(key)
        
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Unable to find appropriate key"
        )
```

#### Authentication Middleware

```python
# app/middleware/auth.py
from fastapi import Request, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from app.core.auth.jwt_handler import JWTHandler

security = HTTPBearer()
jwt_handler = JWTHandler()

async def authenticate_request(request: Request, credentials: HTTPAuthorizationCredentials):
    """Authenticate incoming requests"""
    if not credentials or not credentials.credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Authentication required"
        )
    
    token = credentials.credentials
    
    # Validate Office.js SSO token
    user_info = await jwt_handler.validate_office_token(token)
    
    # Add user info to request state
    request.state.user = {
        'id': user_info.get('oid'),  # Object ID
        'email': user_info.get('preferred_username'),
        'name': user_info.get('name'),
        'tenant_id': user_info.get('tid')
    }
    
    return request.state.user
```

### Security Best Practices Implementation

#### 1. Rate Limiting
```python
# app/middleware/rate_limiting.py
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

limiter = Limiter(key_func=get_remote_address)

# Apply to chat endpoint
@limiter.limit("10/minute")
async def chat_endpoint(request: Request, ...):
    pass
```

#### 2. Input Validation & Sanitization
```python
# app/schemas/chat.py
from pydantic import BaseModel, Field, validator
from typing import List, Optional
import re

class ChatRequest(BaseModel):
    message: str = Field(..., min_length=1, max_length=1000)
    context: Optional[ExcelContext] = None
    timestamp: str = Field(...)
    
    @validator('message')
    def sanitize_message(cls, v):
        # Remove potentially harmful content
        v = re.sub(r'<[^>]*>', '', v)  # Remove HTML tags
        v = v.strip()
        return v

class ExcelContext(BaseModel):
    selectedRange: Optional[SelectedRange] = None
    worksheet: Optional[WorksheetInfo] = None

class SelectedRange(BaseModel):
    address: str = Field(..., regex=r'^[A-Z]+\d+:[A-Z]+\d+$')
    values: List[List[str]] = Field(..., max_items=1000)
    rowCount: int = Field(..., ge=1, le=1000)
    columnCount: int = Field(..., ge=1, le=50)
```

#### 3. Secure Headers Middleware
```python
# app/middleware/security.py
from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        
        # Security headers
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-XSS-Protection"] = "1; mode=block"
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
        response.headers["Content-Security-Policy"] = "default-src 'self'"
        
        return response
```

---

## AI Integration Architecture

### OpenAI Service Implementation

#### Core AI Service
```python
# app/services/ai_service.py
import openai
from typing import Dict, Any, Optional
from app.core.config import settings
from app.core.ai.prompt_templates import PromptTemplates
from app.schemas.chat import ChatRequest, ChatResponse
from app.utils.logging import logger

class AIService:
    def __init__(self):
        self.client = openai.AsyncOpenAI(api_key=settings.OPENAI_API_KEY)
        self.model = settings.OPENAI_MODEL
        self.max_tokens = settings.OPENAI_MAX_TOKENS
        self.prompt_templates = PromptTemplates()
    
    async def generate_response(
        self, 
        request: ChatRequest, 
        user_context: Dict[str, Any]
    ) -> ChatResponse:
        """Generate AI response for chat request"""
        try:
            # Build context-aware prompt
            prompt = await self._build_prompt(request, user_context)
            
            # Call OpenAI API
            completion = await self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": prompt["system"]},
                    {"role": "user", "content": prompt["user"]}
                ],
                max_tokens=self.max_tokens,
                temperature=0.7,
                top_p=0.9
            )
            
            response_text = completion.choices[0].message.content
            
            # Log usage for monitoring
            await self._log_ai_usage(completion.usage, user_context)
            
            return ChatResponse(response=response_text)
            
        except Exception as e:
            logger.error(f"AI service error: {str(e)}")
            raise AIServiceException(f"Failed to generate response: {str(e)}")
    
    async def _build_prompt(
        self, 
        request: ChatRequest, 
        user_context: Dict[str, Any]
    ) -> Dict[str, str]:
        """Build context-aware prompt for AI"""
        
        # Analyze Excel context if available
        excel_analysis = await self._analyze_excel_context(request.context)
        
        # Get appropriate prompt template
        template = self.prompt_templates.get_chat_template(
            has_excel_data=bool(request.context),
            data_type=excel_analysis.get('data_type'),
            user_intent=await self._classify_user_intent(request.message)
        )
        
        return {
            "system": template.format(
                excel_context=excel_analysis.get('summary', ''),
                data_insights=excel_analysis.get('insights', ''),
                user_name=user_context.get('name', 'User')
            ),
            "user": request.message
        }
    
    async def _analyze_excel_context(self, context: Optional[Any]) -> Dict[str, str]:
        """Analyze Excel context data"""
        if not context or not context.selectedRange:
            return {"data_type": "none", "summary": "", "insights": ""}
        
        range_data = context.selectedRange
        values = range_data.values
        
        # Basic data analysis
        analysis = {
            "data_type": self._detect_data_type(values),
            "summary": f"Selected range {range_data.address} contains {range_data.rowCount} rows and {range_data.columnCount} columns",
            "insights": await self._generate_data_insights(values)
        }
        
        return analysis
    
    def _detect_data_type(self, values: List[List[str]]) -> str:
        """Detect the type of data in Excel range"""
        if not values or len(values) < 2:
            return "basic"
        
        # Check for numerical data
        numerical_count = 0
        total_cells = 0
        
        for row in values[1:]:  # Skip header row
            for cell in row:
                total_cells += 1
                try:
                    float(cell.replace(',', '').replace('$', ''))
                    numerical_count += 1
                except (ValueError, AttributeError):
                    pass
        
        if numerical_count / total_cells > 0.7:
            return "numerical"
        elif any(keyword in str(values[0]).lower() for keyword in ['date', 'time', 'month', 'year']):
            return "temporal"
        else:
            return "categorical"
    
    async def _generate_data_insights(self, values: List[List[str]]) -> str:
        """Generate basic insights about the data"""
        if not values or len(values) < 2:
            return "No data available for analysis"
        
        insights = []
        headers = values[0] if values else []
        data_rows = values[1:] if len(values) > 1 else []
        
        insights.append(f"Data contains {len(headers)} columns: {', '.join(headers[:5])}")
        
        if data_rows:
            insights.append(f"Sample data from first row: {', '.join(str(cell) for cell in data_rows[0][:3])}")
        
        return ". ".join(insights)
    
    async def _classify_user_intent(self, message: str) -> str:
        """Classify user intent for better prompt selection"""
        message_lower = message.lower()
        
        if any(word in message_lower for word in ['formula', 'calculate', 'sum', 'average']):
            return "formula_help"
        elif any(word in message_lower for word in ['chart', 'graph', 'visualize', 'plot']):
            return "visualization"
        elif any(word in message_lower for word in ['analyze', 'insight', 'pattern', 'trend']):
            return "analysis"
        elif any(word in message_lower for word in ['format', 'style', 'color', 'appearance']):
            return "formatting"
        else:
            return "general"
    
    async def _log_ai_usage(self, usage: Any, user_context: Dict[str, Any]):
        """Log AI API usage for monitoring and billing"""
        logger.info(
            "AI API Usage",
            extra={
                "user_id": user_context.get('id'),
                "prompt_tokens": usage.prompt_tokens,
                "completion_tokens": usage.completion_tokens,
                "total_tokens": usage.total_tokens,
                "model": self.model
            }
        )
```

#### Prompt Templates System

```python
# app/core/ai/prompt_templates.py
from typing import Dict
from enum import Enum

class UserIntent(Enum):
    FORMULA_HELP = "formula_help"
    VISUALIZATION = "visualization"
    ANALYSIS = "analysis"
    FORMATTING = "formatting"
    GENERAL = "general"

class DataType(Enum):
    NUMERICAL = "numerical"
    CATEGORICAL = "categorical"
    TEMPORAL = "temporal"
    BASIC = "basic"
    NONE = "none"

class PromptTemplates:
    def __init__(self):
        self.base_system_prompt = """
You are an intelligent Excel assistant integrated into Microsoft Excel via an Office.js add-in. 
You help users analyze data, create formulas, generate insights, and provide Excel-specific guidance.

Guidelines:
- Always provide practical, actionable advice specific to Excel
- When suggesting formulas, use proper Excel syntax
- Be concise but thorough in explanations
- Prioritize accuracy and Excel best practices
- Consider the user's current worksheet context when available
"""

        self.context_prompts = {
            DataType.NUMERICAL: """
Current Excel Context: The user has selected numerical data.
Focus on: calculations, statistical analysis, formulas, charts for numerical data.
""",
            DataType.CATEGORICAL: """
Current Excel Context: The user has selected categorical data.
Focus on: data organization, filtering, pivot tables, categorical analysis.
""",
            DataType.TEMPORAL: """
Current Excel Context: The user has selected time-series or date data.
Focus on: date functions, time-based analysis, trend analysis, temporal charts.
""",
            DataType.BASIC: """
Current Excel Context: The user has selected basic data.
Focus on: data organization, basic formulas, formatting, general Excel guidance.
"""
        }

        self.intent_prompts = {
            UserIntent.FORMULA_HELP: """
User Intent: Formula assistance
Provide: Specific Excel formulas, function explanations, formula troubleshooting.
""",
            UserIntent.VISUALIZATION: """
User Intent: Data visualization
Provide: Chart recommendations, visualization best practices, chart creation steps.
""",
            UserIntent.ANALYSIS: """
User Intent: Data analysis
Provide: Analytical insights, statistical summaries, pattern identification.
""",
            UserIntent.FORMATTING: """
User Intent: Formatting help
Provide: Formatting suggestions, styling recommendations, visual improvements.
""",
            UserIntent.GENERAL: """
User Intent: General assistance
Provide: Comprehensive Excel guidance, general best practices.
"""
        }

    def get_chat_template(
        self, 
        has_excel_data: bool = False,
        data_type: str = "none",
        user_intent: str = "general"
    ) -> str:
        """Get appropriate prompt template based on context"""
        
        prompt_parts = [self.base_system_prompt]
        
        # Add context-specific guidance
        if has_excel_data and data_type != "none":
            data_type_enum = DataType(data_type)
            if data_type_enum in self.context_prompts:
                prompt_parts.append(self.context_prompts[data_type_enum])
        
        # Add intent-specific guidance
        intent_enum = UserIntent(user_intent)
        if intent_enum in self.intent_prompts:
            prompt_parts.append(self.intent_prompts[intent_enum])
        
        # Add context placeholders
        if has_excel_data:
            prompt_parts.append("""
Excel Data Context:
{excel_context}

Data Insights:
{data_insights}
""")
        
        prompt_parts.append("""
User: {user_name}
Respond as their helpful Excel assistant, addressing their specific question with the context provided above.
""")
        
        return "\n".join(prompt_parts)
```

### AI Error Handling & Fallbacks

```python
# app/utils/exceptions.py
class AIServiceException(Exception):
    """Custom exception for AI service errors"""
    pass

# app/services/ai_service.py (fallback methods)
class AIService:
    # ... existing code ...
    
    async def get_fallback_response(self, request: ChatRequest) -> ChatResponse:
        """Provide fallback response when AI service fails"""
        message_lower = request.message.lower()
        
        fallback_responses = {
            "formula": "I can help with Excel formulas! Try asking about specific functions like SUM, VLOOKUP, or IF statements.",
            "chart": "For creating charts, select your data and go to Insert > Charts. What type of visualization are you looking for?",
            "data": "I can help analyze your Excel data. Could you select the range you'd like me to examine?",
            "default": "I'm here to help with your Excel tasks. Could you please rephrase your question or try again?"
        }
        
        for keyword, response in fallback_responses.items():
            if keyword in message_lower:
                return ChatResponse(response=response)
        
        return ChatResponse(response=fallback_responses["default"])
```

---

## Data Processing Pipeline

### Excel Data Processing Service

```python
# app/services/excel_service.py
import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
from app.schemas.excel import ExcelContext, DataAnalysis
from app.utils.logging import logger

class ExcelService:
    def __init__(self):
        pass
    
    async def process_excel_context(self, context: ExcelContext) -> DataAnalysis:
        """Process Excel context and generate analysis"""
        if not context or not context.selectedRange:
            return DataAnalysis(
                summary="No data selected",
                data_type="none",
                statistics={},
                insights=[],
                recommendations=[]
            )
        
        try:
            # Convert Excel data to pandas DataFrame
            df = self._create_dataframe(context.selectedRange)
            
            # Perform analysis
            analysis = await self._analyze_dataframe(df, context)
            
            return analysis
            
        except Exception as e:
            logger.error(f"Excel processing error: {str(e)}")
            return DataAnalysis(
                summary=f"Error processing Excel data: {str(e)}",
                data_type="error",
                statistics={},
                insights=[],
                recommendations=[]
            )
    
    def _create_dataframe(self, selected_range) -> pd.DataFrame:
        """Convert Excel range data to pandas DataFrame"""
        values = selected_range.values
        
        if not values or len(values) == 0:
            return pd.DataFrame()
        
        # Assume first row contains headers
        if len(values) > 1:
            headers = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=headers)
        else:
            df = pd.DataFrame(values)
        
        # Clean and convert data types
        df = self._clean_dataframe(df)
        
        return df
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and prepare DataFrame for analysis"""
        # Remove empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Convert numeric columns
        for col in df.columns:
            df[col] = self._convert_to_numeric(df[col])
        
        return df
    
    def _convert_to_numeric(self, series: pd.Series) -> pd.Series:
        """Attempt to convert series to numeric, handling Excel formats"""
        try:
            # Handle common Excel formatting
            if series.dtype == 'object':
                # Remove currency symbols and commas
                cleaned = series.astype(str).str.replace(r'[$,]', '', regex=True)
                # Try to convert to numeric
                numeric_series = pd.to_numeric(cleaned, errors='ignore')
                return numeric_series
            return series
        except Exception:
            return series
    
    async def _analyze_dataframe(self, df: pd.DataFrame, context: ExcelContext) -> DataAnalysis:
        """Perform comprehensive analysis of DataFrame"""
        if df.empty:
            return DataAnalysis(
                summary="Empty dataset",
                data_type="empty",
                statistics={},
                insights=[],
                recommendations=[]
            )
        
        # Basic statistics
        statistics = self._calculate_statistics(df)
        
        # Data type classification
        data_type = self._classify_data_type(df)
        
        # Generate insights
        insights = self._generate_insights(df, statistics)
        
        # Generate recommendations
        recommendations = self._generate_recommendations(df, data_type, insights)
        
        # Create summary
        summary = f"Analyzed {len(df)} rows and {len(df.columns)} columns from range {context.selectedRange.address}"
        
        return DataAnalysis(
            summary=summary,
            data_type=data_type,
            statistics=statistics,
            insights=insights,
            recommendations=recommendations
        )
    
    def _calculate_statistics(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate comprehensive statistics for the DataFrame"""
        stats = {
            "shape": {"rows": len(df), "columns": len(df.columns)},
            "columns": list(df.columns),
            "data_types": df.dtypes.astype(str).to_dict()
        }
        
        # Numeric column statistics
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            stats["numeric_summary"] = df[numeric_cols].describe().to_dict()
        
        # Missing data analysis
        missing_data = df.isnull().sum()
        if missing_data.sum() > 0:
            stats["missing_data"] = missing_data.to_dict()
        
        # Unique values for categorical columns
        categorical_cols = df.select_dtypes(include=['object']).columns
        if len(categorical_cols) > 0:
            stats["categorical_summary"] = {
                col: {"unique_count": df[col].nunique(), "most_common": df[col].mode().iloc[0] if not df[col].mode().empty else None}
                for col in categorical_cols
            }
        
        return stats
    
    def _classify_data_type(self, df: pd.DataFrame) -> str:
        """Classify the overall data type of the DataFrame"""
        numeric_ratio = len(df.select_dtypes(include=[np.number]).columns) / len(df.columns)
        
        if numeric_ratio > 0.7:
            return "numerical"
        elif any(col.lower() in ['date', 'time', 'month', 'year'] for col in df.columns):
            return "temporal"
        elif numeric_ratio > 0.3:
            return "mixed"
        else:
            return "categorical"
    
    def _generate_insights(self, df: pd.DataFrame, statistics: Dict[str, Any]) -> List[str]:
        """Generate data insights from analysis"""
        insights = []
        
        # Data quality insights
        if "missing_data" in statistics:
            missing_ratio = sum(statistics["missing_data"].values()) / (len(df) * len(df.columns))
            if missing_ratio > 0.1:
                insights.append(f"Data quality concern: {missing_ratio:.1%} of cells are empty")
        
        # Numeric insights
        if "numeric_summary" in statistics:
            for col, stats in statistics["numeric_summary"].items():
                if stats.get("std", 0) > stats.get("mean", 0):
                    insights.append(f"Column '{col}' shows high variability")
        
        # Pattern insights
        if len(df) > 10:
            insights.append(f"Dataset contains {len(df)} records, suitable for analysis")
        
        return insights
    
    def _generate_recommendations(self, df: pd.DataFrame, data_type: str, insights: List[str]) -> List[str]:
        """Generate actionable recommendations"""
        recommendations = []
        
        # Data type specific recommendations
        if data_type == "numerical":
            recommendations.extend([
                "Consider creating charts to visualize trends",
                "Use statistical functions like AVERAGE, STDEV, and CORREL",
                "Look for outliers using conditional formatting"
            ])
        elif data_type == "categorical":
            recommendations.extend([
                "Create pivot tables to summarize categories",
                "Use COUNTIF functions for category analysis",
                "Consider creating pie charts for category distribution"
            ])
        elif data_type == "temporal":
            recommendations.extend([
                "Use date functions for time-based analysis",
                "Create time series charts to show trends",
                "Consider using FORECAST functions for predictions"
            ])
        
        # General recommendations
        if len(df.columns) > 5:
            recommendations.append("Consider splitting analysis across multiple worksheets")
        
        if len(df) > 1000:
            recommendations.append("Use filters and sorting to focus on specific data subsets")
        
        return recommendations
```

### Data Schemas

```python
# app/schemas/excel.py
from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional

class SelectedRange(BaseModel):
    address: str = Field(..., description="Excel range address (e.g., A1:B10)")
    values: List[List[str]] = Field(..., description="Cell values as 2D array")
    rowCount: int = Field(..., ge=1, description="Number of rows")
    columnCount: int = Field(..., ge=1, description="Number of columns")

class WorksheetInfo(BaseModel):
    name: str = Field(..., description="Worksheet name")

class ExcelContext(BaseModel):
    selectedRange: Optional[SelectedRange] = None
    worksheet: Optional[WorksheetInfo] = None

class DataAnalysis(BaseModel):
    summary: str = Field(..., description="Analysis summary")
    data_type: str = Field(..., description="Classified data type")
    statistics: Dict[str, Any] = Field(..., description="Statistical analysis")
    insights: List[str] = Field(..., description="Data insights")
    recommendations: List[str] = Field(..., description="Actionable recommendations")
```

---

## API Design & Contracts

### Chat API Endpoint

```python
# app/api/v1/endpoints/chat.py
from fastapi import APIRouter, Depends, HTTPException, status, Request
from fastapi.security import HTTPBearer
from app.schemas.chat import ChatRequest, ChatResponse
from app.services.chat_service import ChatService
from app.services.ai_service import AIService
from app.services.excel_service import ExcelService
from app.middleware.auth import authenticate_request
from app.utils.logging import logger
import time

router = APIRouter()
security = HTTPBearer()

chat_service = ChatService()
ai_service = AIService()
excel_service = ExcelService()

@router.post("/chat", response_model=ChatResponse)
async def chat_endpoint(
    request: Request,
    chat_request: ChatRequest,
    user = Depends(authenticate_request)
):
    """
    Process chat message with Excel context and return AI response.
    
    - **message**: User's question or request
    - **context**: Excel context data (selected range, worksheet info)
    - **timestamp**: Request timestamp for tracking
    """
    start_time = time.time()
    
    try:
        logger.info(
            "Chat request received",
            extra={
                "user_id": user.get('id'),
                "message_length": len(chat_request.message),
                "has_context": bool(chat_request.context)
            }
        )
        
        # Process Excel context if provided
        excel_analysis = None
        if chat_request.context:
            excel_analysis = await excel_service.process_excel_context(chat_request.context)
        
        # Generate AI response
        try:
            ai_response = await ai_service.generate_response(chat_request, user)
        except Exception as e:
            logger.warning(f"AI service failed, using fallback: {str(e)}")
            ai_response = await ai_service.get_fallback_response(chat_request)
        
        # Store conversation history
        await chat_service.store_conversation(
            user_id=user.get('id'),
            message=chat_request.message,
            response=ai_response.response,
            excel_context=chat_request.context,
            excel_analysis=excel_analysis
        )
        
        # Log response metrics
        response_time = time.time() - start_time
        logger.info(
            "Chat response generated",
            extra={
                "user_id": user.get('id'),
                "response_time": response_time,
                "response_length": len(ai_response.response)
            }
        )
        
        return ai_response
        
    except Exception as e:
        logger.error(
            "Chat endpoint error",
            extra={
                "user_id": user.get('id'),
                "error": str(e),
                "request_time": time.time() - start_time
            }
        )
        
        # Return user-friendly error
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Sorry, I'm experiencing technical difficulties. Please try again."
        )

@router.get("/chat/history")
async def get_chat_history(
    user = Depends(authenticate_request),
    limit: int = 20,
    offset: int = 0
):
    """Get user's chat history"""
    try:
        history = await chat_service.get_user_history(
            user_id=user.get('id'),
            limit=limit,
            offset=offset
        )
        return {"conversations": history, "total": len(history)}
        
    except Exception as e:
        logger.error(f"Error fetching chat history: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Unable to fetch chat history"
        )

@router.delete("/chat/history")
async def clear_chat_history(user = Depends(authenticate_request)):
    """Clear user's chat history"""
    try:
        await chat_service.clear_user_history(user_id=user.get('id'))
        return {"message": "Chat history cleared successfully"}
        
    except Exception as e:
        logger.error(f"Error clearing chat history: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Unable to clear chat history"
        )
```

### Health Check & Monitoring Endpoints

```python
# app/api/v1/endpoints/health.py
from fastapi import APIRouter, Depends
from app.services.health_service import HealthService
from app.schemas.health import HealthResponse, DetailedHealthResponse

router = APIRouter()
health_service = HealthService()

@router.get("/health", response_model=HealthResponse)
async def health_check():
    """Basic health check endpoint"""
    return await health_service.basic_health_check()

@router.get("/health/detailed", response_model=DetailedHealthResponse)
async def detailed_health_check():
    """Detailed health check including external dependencies"""
    return await health_service.detailed_health_check()

@router.get("/metrics")
async def get_metrics():
    """Prometheus-compatible metrics endpoint"""
    return await health_service.get_prometheus_metrics()
```

### API Response Models

```python
# app/schemas/chat.py
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any
from datetime import datetime

class ChatRequest(BaseModel):
    message: str = Field(..., min_length=1, max_length=2000, description="User's message")
    context: Optional[ExcelContext] = Field(None, description="Excel context data")
    timestamp: str = Field(..., description="Request timestamp")

class ChatResponse(BaseModel):
    response: str = Field(..., description="AI-generated response")
    metadata: Optional[Dict[str, Any]] = Field(None, description="Response metadata")

class ConversationHistory(BaseModel):
    id: str = Field(..., description="Conversation ID")
    user_message: str = Field(..., description="User's message")
    ai_response: str = Field(..., description="AI response")
    timestamp: datetime = Field(..., description="Conversation timestamp")
    excel_context: Optional[ExcelContext] = Field(None, description="Associated Excel context")

# app/schemas/health.py
class HealthResponse(BaseModel):
    status: str = Field(..., description="Health status: healthy, degraded, unhealthy")
    timestamp: datetime = Field(..., description="Health check timestamp")
    version: str = Field(..., description="Application version")

class ServiceHealth(BaseModel):
    name: str
    status: str
    response_time_ms: Optional[float] = None
    error: Optional[str] = None

class DetailedHealthResponse(BaseModel):
    status: str
    timestamp: datetime
    version: str
    services: List[ServiceHealth]
    system_info: Dict[str, Any]
```

---

## Infrastructure & Deployment

### Docker Configuration

```dockerfile
# deployment/docker/Dockerfile
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Create app user
RUN useradd --create-home --shell /bin/bash appuser

# Set work directory
WORKDIR /app

# Copy requirements
COPY requirements/prod.txt requirements.txt

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Change ownership to app user
RUN chown -R appuser:appuser /app
USER appuser

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import requests; requests.get('http://localhost:8000/health')"

# Expose port
EXPOSE 8000

# Run application
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
```

```yaml
# deployment/docker/docker-compose.yml
version: '3.8'

services:
  app:
    build:
      context: ../..
      dockerfile: deployment/docker/Dockerfile
    ports:
      - "8000:8000"
    environment:
      - DATABASE_URL=postgresql://postgres:password@db:5432/excel_ai
      - REDIS_URL=redis://redis:6379
      - OPENAI_API_KEY=${OPENAI_API_KEY}
      - SECRET_KEY=${SECRET_KEY}
    depends_on:
      - db
      - redis
    volumes:
      - app_logs:/app/logs

  db:
    image: postgres:15-alpine
    environment:
      - POSTGRES_DB=excel_ai
      - POSTGRES_USER=postgres
      - POSTGRES_PASSWORD=password
    volumes:
      - postgres_data:/var/lib/postgresql/data
    ports:
      - "5432:5432"

  redis:
    image: redis:7-alpine
    ports:
      - "6379:6379"
    volumes:
      - redis_data:/data

  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf
      - ./ssl:/etc/nginx/ssl
    depends_on:
      - app

volumes:
  postgres_data:
  redis_data:
  app_logs:
```

### Azure Container Apps Deployment

```yaml
# deployment/azure/container-apps.yaml
apiVersion: 2021-03-01
name: excel-ai-backend
location: East US
identity:
  type: SystemAssigned
properties:
  managedEnvironmentId: /subscriptions/{subscription-id}/resourceGroups/{rg}/providers/Microsoft.App/managedEnvironments/{env-name}
  configuration:
    secrets:
      - name: openai-api-key
        value: ${OPENAI_API_KEY}
      - name: database-url
        value: ${DATABASE_URL}
      - name: secret-key
        value: ${SECRET_KEY}
    ingress:
      external: true
      targetPort: 8000
      allowInsecure: false
      traffic:
        - weight: 100
          latestRevision: true
    dapr:
      enabled: false
  template:
    containers:
      - image: your-registry.azurecr.io/excel-ai-backend:latest
        name: excel-ai-backend
        env:
          - name: OPENAI_API_KEY
            secretRef: openai-api-key
          - name: DATABASE_URL
            secretRef: database-url
          - name: SECRET_KEY
            secretRef: secret-key
          - name: ENVIRONMENT
            value: production
        resources:
          cpu: 1.0
          memory: 2.0Gi
        probes:
          - type: liveness
            httpGet:
              path: /health
              port: 8000
            initialDelaySeconds: 30
            periodSeconds: 10
          - type: readiness
            httpGet:
              path: /health
              port: 8000
            initialDelaySeconds: 5
            periodSeconds: 5
    scale:
      minReplicas: 1
      maxReplicas: 10
      rules:
        - name: http-rule
          http:
            metadata:
              concurrentRequests: '10'
```

### CI/CD Pipeline

```yaml
# .github/workflows/deploy.yml
name: Deploy Excel AI Backend

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

env:
  REGISTRY: your-registry.azurecr.io
  IMAGE_NAME: excel-ai-backend

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.11'
          
      - name: Install dependencies
        run: |
          pip install -r requirements/dev.txt
          
      - name: Run tests
        run: |
          pytest tests/ --cov=app --cov-report=xml
          
      - name: Upload coverage
        uses: codecov/codecov-action@v3

  build:
    needs: test
    runs-on: ubuntu-latest
    if: github.ref == 'refs/heads/main'
    
    steps:
      - uses: actions/checkout@v4
      
      - name: Log in to registry
        uses: azure/docker-login@v1
        with:
          login-server: ${{ env.REGISTRY }}
          username: ${{ secrets.REGISTRY_USERNAME }}
          password: ${{ secrets.REGISTRY_PASSWORD }}
          
      - name: Build and push image
        run: |
          docker build -f deployment/docker/Dockerfile -t ${{ env.REGISTRY }}/${{ env.IMAGE_NAME }}:${{ github.sha }} .
          docker push ${{ env.REGISTRY }}/${{ env.IMAGE_NAME }}:${{ github.sha }}

  deploy:
    needs: build
    runs-on: ubuntu-latest
    environment: production
    
    steps:
      - uses: actions/checkout@v4
      
      - name: Azure Login
        uses: azure/login@v1
        with:
          creds: ${{ secrets.AZURE_CREDENTIALS }}
          
      - name: Deploy to Container Apps
        run: |
          az containerapp update \
            --name excel-ai-backend \
            --resource-group excel-ai-rg \
            --image ${{ env.REGISTRY }}/${{ env.IMAGE_NAME }}:${{ github.sha }}
```

---

## Development Environment Setup

### Requirements Files

```txt
# requirements/base.txt
fastapi==0.104.1
uvicorn[standard]==0.24.0
pydantic==2.5.0
pydantic-settings==2.1.0
sqlalchemy==2.0.23
alembic==1.12.1
asyncpg==0.29.0
redis==5.0.1
openai==1.3.7
pandas==2.1.4
numpy==1.25.2
python-jose[cryptography]==3.3.0
python-multipart==0.0.6
httpx==0.25.2
```

```txt
# requirements/dev.txt
-r base.txt
pytest==7.4.3
pytest-asyncio==0.21.1
pytest-cov==4.1.0
pytest-mock==3.12.0
black==23.11.0
isort==5.12.0
flake8==6.1.0
mypy==1.7.1
pre-commit==3.6.0
```

```txt
# requirements/prod.txt
-r base.txt
gunicorn==21.2.0
prometheus-client==0.19.0
structlog==23.2.0
```

### Development Scripts

```python
# scripts/start_dev.py
#!/usr/bin/env python3
"""
Development server startup script
"""
import subprocess
import sys
import os
from pathlib import Path

def check_dependencies():
    """Check if all dependencies are installed"""
    try:
        import uvicorn
        import fastapi
        print("✓ Dependencies are installed")
    except ImportError as e:
        print(f"✗ Missing dependency: {e}")
        print("Run: pip install -r requirements/dev.txt")
        sys.exit(1)

def check_env_file():
    """Check if .env file exists"""
    env_file = Path(".env")
    if not env_file.exists():
        print("✗ .env file not found")
        print("Copy .env.example to .env and configure your settings")
        sys.exit(1)
    print("✓ Environment file found")

def start_server():
    """Start the development server"""
    print("🚀 Starting Excel AI Backend development server...")
    
    cmd = [
        "uvicorn",
        "app.main:app",
        "--reload",
        "--host", "0.0.0.0",
        "--port", "8000",
        "--log-level", "debug"
    ]
    
    subprocess.run(cmd)

if __name__ == "__main__":
    check_dependencies()
    check_env_file()
    start_server()
```

```python
# scripts/migrate_db.py
#!/usr/bin/env python3
"""
Database migration script
"""
import subprocess
import sys

def run_migrations():
    """Run Alembic database migrations"""
    print("Running database migrations...")
    
    commands = [
        ["alembic", "upgrade", "head"],
    ]
    
    for cmd in commands:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"✗ Command failed: {' '.join(cmd)}")
            print(f"Error: {result.stderr}")
            sys.exit(1)
        else:
            print(f"✓ {' '.join(cmd)}")

if __name__ == "__main__":
    run_migrations()
```

### Environment Configuration

```bash
# .env.example
# Application
APP_NAME=Excel AI Chat Backend
DEBUG=true
SECRET_KEY=your-secret-key-here
API_V1_STR=/api/v1

# Database
DATABASE_URL=postgresql://user:password@localhost:5432/excel_ai

# Redis
REDIS_URL=redis://localhost:6379

# OpenAI
OPENAI_API_KEY=your-openai-api-key
OPENAI_MODEL=gpt-4
OPENAI_MAX_TOKENS=1000

# Microsoft Azure
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret

# CORS
ALLOWED_ORIGINS=https://mavencat.github.io,https://localhost:3000

# Logging
LOG_LEVEL=DEBUG

# Development
ENVIRONMENT=development
```

---

## Testing Strategy

### Test Structure

```python
# tests/conftest.py
import pytest
import asyncio
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from app.main import app
from app.models.base import Base
from app.core.config import settings

# Test database
SQLALCHEMY_DATABASE_URL = "sqlite:///./test.db"
engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})
TestingSessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

@pytest.fixture(scope="session")
def event_loop():
    """Create an instance of the default event loop for the test session."""
    loop = asyncio.get_event_loop_policy().new_event_loop()
    yield loop
    loop.close()

@pytest.fixture
def client():
    """Create test client"""
    Base.metadata.create_all(bind=engine)
    with TestClient(app) as c:
        yield c
    Base.metadata.drop_all(bind=engine)

@pytest.fixture
def mock_openai_response():
    """Mock OpenAI API response"""
    return {
        "choices": [{
            "message": {
                "content": "This is a test AI response about your Excel data."
            }
        }],
        "usage": {
            "prompt_tokens": 50,
            "completion_tokens": 30,
            "total_tokens": 80
        }
    }

@pytest.fixture
def excel_context():
    """Sample Excel context for testing"""
    return {
        "selectedRange": {
            "address": "A1:B3",
            "values": [
                ["Name", "Value"],
                ["Item 1", "100"],
                ["Item 2", "200"]
            ],
            "rowCount": 3,
            "columnCount": 2
        },
        "worksheet": {
            "name": "Sheet1"
        }
    }

@pytest.fixture
def auth_headers():
    """Mock authentication headers"""
    return {"Authorization": "Bearer mock-jwt-token"}
```

### Unit Tests

```python
# tests/unit/test_excel_service.py
import pytest
from app.services.excel_service import ExcelService
from app.schemas.excel import ExcelContext, SelectedRange, WorksheetInfo

class TestExcelService:
    
    @pytest.fixture
    def excel_service(self):
        return ExcelService()
    
    @pytest.mark.asyncio
    async def test_process_excel_context_with_valid_data(self, excel_service, excel_context):
        """Test Excel context processing with valid data"""
        context = ExcelContext(**excel_context)
        
        result = await excel_service.process_excel_context(context)
        
        assert result.summary is not None
        assert result.data_type in ["numerical", "categorical", "mixed", "temporal"]
        assert isinstance(result.statistics, dict)
        assert isinstance(result.insights, list)
        assert isinstance(result.recommendations, list)
    
    @pytest.mark.asyncio
    async def test_process_excel_context_empty_data(self, excel_service):
        """Test Excel context processing with empty data"""
        result = await excel_service.process_excel_context(None)
        
        assert result.summary == "No data selected"
        assert result.data_type == "none"
        assert result.statistics == {}
        assert result.insights == []
        assert result.recommendations == []
    
    def test_create_dataframe(self, excel_service):
        """Test DataFrame creation from Excel data"""
        selected_range = SelectedRange(
            address="A1:B3",
            values=[["Name", "Value"], ["Item 1", "100"], ["Item 2", "200"]],
            rowCount=3,
            columnCount=2
        )
        
        df = excel_service._create_dataframe(selected_range)
        
        assert len(df) == 2  # Excluding header
        assert list(df.columns) == ["Name", "Value"]
        assert df["Name"].iloc[0] == "Item 1"
```

### Integration Tests

```python
# tests/integration/test_chat_api.py
import pytest
from unittest.mock import AsyncMock, patch
from fastapi.testclient import TestClient

class TestChatAPI:
    
    @patch('app.services.ai_service.AIService.generate_response')
    @patch('app.middleware.auth.authenticate_request')
    def test_chat_endpoint_success(self, mock_auth, mock_ai, client, excel_context, mock_openai_response):
        """Test successful chat endpoint call"""
        # Mock authentication
        mock_auth.return_value = {"id": "test-user", "email": "test@example.com"}
        
        # Mock AI service
        mock_ai.return_value = AsyncMock()
        mock_ai.return_value.response = "This is a test response"
        
        # Test request
        response = client.post(
            "/api/v1/chat",
            json={
                "message": "What data do I have selected?",
                "context": excel_context,
                "timestamp": "2024-01-01T00:00:00Z"
            },
            headers={"Authorization": "Bearer test-token"}
        )
        
        assert response.status_code == 200
        data = response.json()
        assert "response" in data
        assert isinstance(data["response"], str)
    
    def test_chat_endpoint_unauthorized(self, client):
        """Test chat endpoint without authentication"""
        response = client.post(
            "/api/v1/chat",
            json={
                "message": "Test message",
                "timestamp": "2024-01-01T00:00:00Z"
            }
        )
        
        assert response.status_code == 401
    
    def test_chat_endpoint_invalid_data(self, client):
        """Test chat endpoint with invalid data"""
        response = client.post(
            "/api/v1/chat",
            json={
                "message": "",  # Empty message
                "timestamp": "2024-01-01T00:00:00Z"
            },
            headers={"Authorization": "Bearer test-token"}
        )
        
        assert response.status_code == 422  # Validation error
```

### End-to-End Tests

```python
# tests/e2e/test_full_flow.py
import pytest
from unittest.mock import patch, AsyncMock

class TestFullFlow:
    
    @patch('app.core.auth.jwt_handler.JWTHandler.validate_office_token')
    @patch('app.services.ai_service.AIService.generate_response')
    def test_complete_chat_flow(self, mock_ai, mock_jwt, client, excel_context):
        """Test complete flow from authentication to AI response"""
        # Mock JWT validation
        mock_jwt.return_value = {
            "oid": "test-user-id",
            "preferred_username": "test@example.com",
            "name": "Test User"
        }
        
        # Mock AI response
        mock_ai.return_value = AsyncMock()
        mock_ai.return_value.response = "Based on your selected data, you have 2 items with a total value of 300."
        
        # Send chat request
        response = client.post(
            "/api/v1/chat",
            json={
                "message": "Analyze my selected data",
                "context": excel_context,
                "timestamp": "2024-01-01T00:00:00Z"
            },
            headers={"Authorization": "Bearer valid-office-token"}
        )
        
        # Verify response
        assert response.status_code == 200
        data = response.json()
        assert "response" in data
        assert "data" in data["response"].lower()
        
        # Verify JWT validation was called
        mock_jwt.assert_called_once()
        
        # Verify AI service was called with correct parameters
        mock_ai.assert_called_once()
```

---

## Monitoring & Observability

### Logging Configuration

```python
# app/utils/logging.py
import structlog
import logging
from typing import Dict, Any
from app.core.config import settings

def configure_logging():
    """Configure structured logging"""
    
    # Configure stdlib logging
    logging.basicConfig(
        format="%(message)s",
        stream=sys.stdout,
        level=getattr(logging, settings.LOG_LEVEL.upper())
    )
    
    # Configure structlog
    structlog.configure(
        processors=[
            structlog.stdlib.filter_by_level,
            structlog.stdlib.add_logger_name,
            structlog.stdlib.add_log_level,
            structlog.stdlib.PositionalArgumentsFormatter(),
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.processors.UnicodeDecoder(),
            structlog.processors.JSONRenderer()
        ],
        context_class=dict,
        logger_factory=structlog.stdlib.LoggerFactory(),
        wrapper_class=structlog.stdlib.BoundLogger,
        cache_logger_on_first_use=True,
    )

def get_logger(name: str = None):
    """Get structured logger instance"""
    return structlog.get_logger(name)

class RequestLoggingMiddleware:
    """Log all incoming requests"""
    
    def __init__(self, app):
        self.app = app
        self.logger = get_logger("request")
    
    async def __call__(self, scope, receive, send):
        if scope["type"] == "http":
            start_time = time.time()
            
            # Log request
            self.logger.info(
                "Request started",
                method=scope["method"],
                path=scope["path"],
                query_string=scope["query_string"].decode(),
                client=scope.get("client", ["unknown", 0])[0]
            )
            
            # Process request
            await self.app(scope, receive, send)
            
            # Log response
            duration = time.time() - start_time
            self.logger.info(
                "Request completed",
                method=scope["method"],
                path=scope["path"],
                duration_ms=round(duration * 1000, 2)
            )
        else:
            await self.app(scope, receive, send)
```

### Metrics Collection

```python
# app/services/metrics_service.py
from prometheus_client import Counter, Histogram, Gauge, generate_latest
from typing import Dict
import time
import psutil
import asyncio

# Prometheus metrics
REQUEST_COUNT = Counter('http_requests_total', 'Total HTTP requests', ['method', 'endpoint', 'status'])
REQUEST_DURATION = Histogram('http_request_duration_seconds', 'HTTP request duration', ['method', 'endpoint'])
AI_REQUESTS = Counter('ai_requests_total', 'Total AI API requests', ['model', 'status'])
AI_TOKENS = Counter('ai_tokens_total', 'Total AI tokens used', ['model', 'type'])
ACTIVE_USERS = Gauge('active_users', 'Number of active users')
SYSTEM_CPU = Gauge('system_cpu_percent', 'System CPU usage percentage')
SYSTEM_MEMORY = Gauge('system_memory_percent', 'System memory usage percentage')

class MetricsService:
    def __init__(self):
        self.active_users_set = set()
    
    def record_request(self, method: str, endpoint: str, status_code: int, duration: float):
        """Record HTTP request metrics"""
        REQUEST_COUNT.labels(method=method, endpoint=endpoint, status=status_code).inc()
        REQUEST_DURATION.labels(method=method, endpoint=endpoint).observe(duration)
    
    def record_ai_request(self, model: str, status: str, prompt_tokens: int = 0, completion_tokens: int = 0):
        """Record AI API usage metrics"""
        AI_REQUESTS.labels(model=model, status=status).inc()
        if prompt_tokens > 0:
            AI_TOKENS.labels(model=model, type='prompt').inc(prompt_tokens)
        if completion_tokens > 0:
            AI_TOKENS.labels(model=model, type='completion').inc(completion_tokens)
    
    def update_active_users(self, user_id: str):
        """Update active users count"""
        self.active_users_set.add(user_id)
        ACTIVE_USERS.set(len(self.active_users_set))
    
    async def collect_system_metrics(self):
        """Collect system metrics periodically"""
        while True:
            cpu_percent = psutil.cpu_percent()
            memory_percent = psutil.virtual_memory().percent
            
            SYSTEM_CPU.set(cpu_percent)
            SYSTEM_MEMORY.set(memory_percent)
            
            # Clear old active users (older than 1 hour)
            # Implementation depends on your session management
            
            await asyncio.sleep(30)  # Collect every 30 seconds
    
    def get_prometheus_metrics(self) -> str:
        """Get Prometheus-formatted metrics"""
        return generate_latest()
```

### Health Check Service

```python
# app/services/health_service.py
import asyncio
import time
from typing import Dict, List
import httpx
import redis
from sqlalchemy import text
from app.core.config import settings
from app.core.database import get_db
from app.schemas.health import HealthResponse, DetailedHealthResponse, ServiceHealth

class HealthService:
    def __init__(self):
        self.redis_client = redis.from_url(settings.REDIS_URL)
    
    async def basic_health_check(self) -> HealthResponse:
        """Basic health check"""
        return HealthResponse(
            status="healthy",
            timestamp=datetime.utcnow(),
            version="1.0.0"
        )
    
    async def detailed_health_check(self) -> DetailedHealthResponse:
        """Detailed health check including external services"""
        services = []
        overall_status = "healthy"
        
        # Database health
        db_health = await self._check_database()
        services.append(db_health)
        if db_health.status != "healthy":
            overall_status = "degraded"
        
        # Redis health
        redis_health = await self._check_redis()
        services.append(redis_health)
        if redis_health.status != "healthy":
            overall_status = "degraded"
        
        # OpenAI API health
        openai_health = await self._check_openai()
        services.append(openai_health)
        if openai_health.status != "healthy":
            overall_status = "degraded"
        
        # System info
        system_info = await self._get_system_info()
        
        return DetailedHealthResponse(
            status=overall_status,
            timestamp=datetime.utcnow(),
            version="1.0.0",
            services=services,
            system_info=system_info
        )
    
    async def _check_database(self) -> ServiceHealth:
        """Check database connectivity"""
        start_time = time.time()
        try:
            async with get_db() as db:
                await db.execute(text("SELECT 1"))
            response_time = (time.time() - start_time) * 1000
            return ServiceHealth(
                name="database",
                status="healthy",
                response_time_ms=response_time
            )
        except Exception as e:
            return ServiceHealth(
                name="database",
                status="unhealthy",
                error=str(e)
            )
    
    async def _check_redis(self) -> ServiceHealth:
        """Check Redis connectivity"""
        start_time = time.time()
        try:
            await self.redis_client.ping()
            response_time = (time.time() - start_time) * 1000
            return ServiceHealth(
                name="redis",
                status="healthy",
                response_time_ms=response_time
            )
        except Exception as e:
            return ServiceHealth(
                name="redis",
                status="unhealthy",
                error=str(e)
            )
    
    async def _check_openai(self) -> ServiceHealth:
        """Check OpenAI API connectivity"""
        start_time = time.time()
        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    "https://api.openai.com/v1/models",
                    headers={"Authorization": f"Bearer {settings.OPENAI_API_KEY}"},
                    timeout=5.0
                )
                response.raise_for_status()
            response_time = (time.time() - start_time) * 1000
            return ServiceHealth(
                name="openai",
                status="healthy",
                response_time_ms=response_time
            )
        except Exception as e:
            return ServiceHealth(
                name="openai",
                status="unhealthy",
                error=str(e)
            )
    
    async def _get_system_info(self) -> Dict[str, Any]:
        """Get system information"""
        import psutil
        return {
            "cpu_percent": psutil.cpu_percent(),
            "memory_percent": psutil.virtual_memory().percent,
            "disk_percent": psutil.disk_usage('/').percent,
            "python_version": sys.version
        }
```

---

## Security Considerations

### Data Protection & Privacy

#### Sensitive Data Handling
```python
# app/utils/security.py
import hashlib
import hmac
from cryptography.fernet import Fernet
from app.core.config import settings

class DataProtection:
    def __init__(self):
        self.cipher_suite = Fernet(settings.DATA_ENCRYPTION_KEY.encode())
    
    def encrypt_sensitive_data(self, data: str) -> str:
        """Encrypt sensitive data before storage"""
        return self.cipher_suite.encrypt(data.encode()).decode()
    
    def decrypt_sensitive_data(self, encrypted_data: str) -> str:
        """Decrypt sensitive data for processing"""
        return self.cipher_suite.decrypt(encrypted_data.encode()).decode()
    
    def hash_user_id(self, user_id: str) -> str:
        """Create non-reversible hash of user ID"""
        return hashlib.sha256(f"{user_id}{settings.SECRET_KEY}".encode()).hexdigest()
    
    def sanitize_excel_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Remove potentially sensitive information from Excel data"""
        # Remove email patterns
        # Remove phone number patterns
        # Remove SSN patterns
        # etc.
        return data
```

#### Data Retention Policy
```python
# app/services/data_retention_service.py
from datetime import datetime, timedelta
from app.models.chat import Conversation
from app.core.database import get_db

class DataRetentionService:
    def __init__(self):
        self.retention_days = 90  # 90 days retention
    
    async def cleanup_old_conversations(self):
        """Remove conversations older than retention period"""
        cutoff_date = datetime.utcnow() - timedelta(days=self.retention_days)
        
        async with get_db() as db:
            await db.query(Conversation).filter(
                Conversation.created_at < cutoff_date
            ).delete()
            await db.commit()
    
    async def anonymize_user_data(self, user_id: str):
        """Anonymize user data for GDPR compliance"""
        async with get_db() as db:
            conversations = await db.query(Conversation).filter(
                Conversation.user_id == user_id
            ).all()
            
            for conv in conversations:
                conv.user_id = "anonymized"
                conv.user_message = "[REDACTED]"
            
            await db.commit()
```

### Security Headers & Middleware

```python
# app/middleware/security.py
from fastapi import Request, Response
from starlette.middleware.base import BaseHTTPMiddleware
import secrets

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        
        # Generate nonce for CSP
        nonce = secrets.token_urlsafe(16)
        
        # Security headers
        response.headers.update({
            "X-Content-Type-Options": "nosniff",
            "X-Frame-Options": "DENY",
            "X-XSS-Protection": "1; mode=block",
            "Strict-Transport-Security": "max-age=31536000; includeSubDomains; preload",
            "Referrer-Policy": "strict-origin-when-cross-origin",
            "Permissions-Policy": "geolocation=(), microphone=(), camera=()",
            "Content-Security-Policy": f"default-src 'self'; script-src 'self' 'nonce-{nonce}'; style-src 'self' 'unsafe-inline'",
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        })
        
        return response
```

---

## Performance Optimization

### Caching Strategy

```python
# app/services/cache_service.py
import json
import redis
from typing import Any, Optional
from app.core.config import settings

class CacheService:
    def __init__(self):
        self.redis_client = redis.from_url(settings.REDIS_URL)
        self.default_ttl = 3600  # 1 hour
    
    async def get(self, key: str) -> Optional[Any]:
        """Get value from cache"""
        try:
            value = await self.redis_client.get(key)
            return json.loads(value) if value else None
        except Exception:
            return None
    
    async def set(self, key: str, value: Any, ttl: int = None) -> bool:
        """Set value in cache"""
        try:
            ttl = ttl or self.default_ttl
            serialized = json.dumps(value, default=str)
            await self.redis_client.setex(key, ttl, serialized)
            return True
        except Exception:
            return False
    
    async def delete(self, key: str) -> bool:
        """Delete value from cache"""
        try:
            await self.redis_client.delete(key)
            return True
        except Exception:
            return False
    
    def get_cache_key(self, prefix: str, *args) -> str:
        """Generate cache key"""
        return f"{prefix}:{':'.join(str(arg) for arg in args)}"

# Usage in services
class AIService:
    def __init__(self):
        self.cache = CacheService()
    
    async def generate_response(self, request: ChatRequest, user_context: Dict[str, Any]) -> ChatResponse:
        # Try cache first for repeated questions
        cache_key = self.cache.get_cache_key(
            "ai_response", 
            user_context.get('id'), 
            request.message[:100]  # First 100 chars
        )
        
        cached_response = await self.cache.get(cache_key)
        if cached_response:
            return ChatResponse(**cached_response)
        
        # Generate new response
        response = await self._generate_new_response(request, user_context)
        
        # Cache for 30 minutes
        await self.cache.set(cache_key, response.dict(), ttl=1800)
        
        return response
```

### Database Optimization

```python
# app/models/base.py
from sqlalchemy import Index
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()

# app/models/chat.py
from sqlalchemy import Column, String, Text, DateTime, ForeignKey, Index
from sqlalchemy.dialects.postgresql import UUID
import uuid

class Conversation(Base):
    __tablename__ = "conversations"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(String(255), nullable=False, index=True)
    user_message = Column(Text, nullable=False)
    ai_response = Column(Text, nullable=False)
    excel_context = Column(Text, nullable=True)  # JSON stored as text
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    response_time_ms = Column(Integer, nullable=True)
    
    # Composite indexes for performance
    __table_args__ = (
        Index('idx_user_created', 'user_id', 'created_at'),
        Index('idx_created_desc', 'created_at', postgresql_using='btree'),
    )
```

### Response Compression

```python
# app/middleware/compression.py
from fastapi import Request, Response
from starlette.middleware.base import BaseHTTPMiddleware
import gzip

class CompressionMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        
        # Check if client accepts gzip
        accept_encoding = request.headers.get("accept-encoding", "")
        if "gzip" in accept_encoding and len(response.body) > 1000:
            # Compress response
            compressed_body = gzip.compress(response.body)
            response.headers["content-encoding"] = "gzip"
            response.headers["content-length"] = str(len(compressed_body))
            response.body = compressed_body
        
        return response
```

---

## Future Enhancements

### Phase 1: Core Functionality (MVP)
- ✅ Basic chat interface integration
- ✅ Excel context processing
- ✅ OpenAI integration
- ✅ Authentication & security
- ✅ Basic deployment

### Phase 2: Enhanced Features (Month 2-3)
- **Advanced Excel Operations**
  - Formula generation and validation
  - Chart creation suggestions
  - Data transformation recommendations
  
- **Conversation Memory**
  - Multi-turn conversation context
  - Conversation history with search
  - User preferences learning

- **Enhanced AI Capabilities**
  - Function calling for Excel operations
  - Multi-modal support (image analysis)
  - Custom prompt templates per user

### Phase 3: Enterprise Features (Month 4-6)
- **Collaboration Features**
  - Team workspaces
  - Shared conversation history
  - Admin dashboard and controls

- **Advanced Analytics**
  - Usage analytics and insights
  - Performance optimization recommendations
  - Custom report generation

- **Integration Expansion**
  - Microsoft Graph integration
  - SharePoint connectivity
  - Power BI integration

### Phase 4: Scale & Innovation (Month 6+)
- **AI Model Fine-tuning**
  - Custom model training on user data
  - Domain-specific Excel expertise
  - Predictive analytics capabilities

- **Multi-Platform Support**
  - Word and PowerPoint add-ins
  - Mobile application
  - Web-based standalone version

- **Advanced Security**
  - Enterprise SSO providers
  - Audit logging and compliance
  - Data sovereignty options

---

## Conclusion

This comprehensive architecture document provides a complete blueprint for developing a production-ready Python backend service for the Excel AI Chat add-in. The design emphasizes:

- **Scalability**: Microservices architecture with clear separation of concerns
- **Security**: Enterprise-grade authentication and data protection
- **Performance**: Caching, optimization, and monitoring strategies
- **Maintainability**: Clean code structure and comprehensive testing
- **Observability**: Detailed logging, metrics, and health checks

The modular design allows for incremental development and deployment, with clear pathways for future enhancements and enterprise features.

**Next Steps for Development Team:**

1. **Environment Setup**: Configure development environment using provided scripts
2. **Core Implementation**: Start with basic FastAPI structure and chat endpoint
3. **Authentication Integration**: Implement Office.js SSO validation
4. **AI Service Development**: Build OpenAI integration with prompt templates
5. **Testing & Deployment**: Implement testing strategy and deploy to staging
6. **Production Deployment**: Configure production environment and monitoring
7. **Iteration & Enhancement**: Add features based on user feedback and requirements

This architecture serves as a comprehensive guide for building a robust, scalable, and secure backend service that seamlessly integrates with the existing Excel Office.js frontend to provide an exceptional AI-powered user experience.

---

**Document Version**: 1.0  
**Last Updated**: December 2024  
**Next Review**: January 2025