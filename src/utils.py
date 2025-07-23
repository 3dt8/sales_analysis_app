import logging

# Cấu hình logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_error(message):
    """
    Ghi lỗi vào file log.
    
    Args:
        message: Thông điệp lỗi
    """
    logging.error(message)