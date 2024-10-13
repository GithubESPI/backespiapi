import httpx
from fastapi import HTTPException
import logging
import requests
from requests.exceptions import RequestException

# Configure the logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)  # Set to INFO or ERROR for production

# Configuration constants
DOCUMENTS_API_URL = 'https://bulletin.groupe-espi.fr/api/documents'  # Replace with the correct base URL
HTTP_TIMEOUT = 60.0

# Function to save the Excel file URL to the database
def save_generated_excel_url_to_db(user_id, excel_url):
    try:
        response = requests.post(
            DOCUMENTS_API_URL,
            json={
                'userId': user_id,
                'generatedExcelUrl': excel_url
            },
            timeout=HTTP_TIMEOUT
        )
        response.raise_for_status()
    except RequestException as e:
        logger.error(f"Failed to save Excel URL: {str(e)}")
        raise Exception(f"Failed to save Excel URL: {str(e)}")

# Asynchronous function to fetch data from an API
async def fetch_api_data(url: str, headers: dict):
    logger.debug(f"Fetching data from {url} with headers {headers}")
    
    async with httpx.AsyncClient(follow_redirects=True) as client:
        try:
            response = await client.get(url, headers=headers, timeout=HTTP_TIMEOUT)
            response.raise_for_status()
        except httpx.RequestError as exc:
            logger.error(f"An error occurred while requesting {exc.request.url!r}: {str(exc)}")
            raise HTTPException(status_code=500, detail="Internal Server Error")
        except httpx.HTTPStatusError as exc:
            logger.error(f"Error response {exc.response.status_code} while requesting {exc.request.url!r}: {exc.response.text}")
            raise HTTPException(status_code=exc.response.status_code, detail=f"API call failed with status {exc.response.status_code}")

        try:
            data = response.json()
            logger.debug(f"Fetched data: {data}")
            if isinstance(data, (list, dict)):
                return data
            else:
                logger.error("Data is not a list or dict")
                raise HTTPException(status_code=500, detail="Invalid data format")
        except ValueError as e:
            logger.error(f"Error parsing JSON: {str(e)}")
            raise HTTPException(status_code=500, detail="Error parsing JSON")
