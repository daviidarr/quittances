import os
import configparser
from main import validate_email_params, validate_config

def test_email_validation():
    # Test valid email params
    try:
        validate_email_params("test@example.com", "password123", "recipient@example.com")
        print("✓ Valid email parameters test passed")
    except ValueError:
        print("✗ Valid email parameters test failed")

    # Test missing params
    try:
        validate_email_params("", "password123", "recipient@example.com")
        print("✗ Missing email parameter test failed")
    except ValueError:
        print("✓ Missing email parameter test passed")

    # Test invalid email format
    try:
        validate_email_params("invalid-email", "password123", "recipient@example.com")
        print("✗ Invalid email format test failed")
    except ValueError:
        print("✓ Invalid email format test passed")

def test_config_validation():
    config = configparser.ConfigParser(interpolation=None)
    
    # Test missing gmail section
    config['property1'] = {
        'address': 'test',
        'tenantname': 'test',
        'tenant_email': 'test@example.com',
        'landlordname': 'test',
        'hors_charge': '100',
        'charge': '50',
        'total_litteral': 'one hundred fifty'
    }
    try:
        validate_config(config)
        print("✗ Missing gmail section test failed")
    except ValueError:
        print("✓ Missing gmail section test passed")

    # Test valid config
    config['gmail'] = {
        'sender_email': 'test@example.com',
        'sender_password': 'password123'
    }
    try:
        validate_config(config)
        print("✓ Valid config test passed")
    except ValueError as e:
        print(f"✗ Valid config test failed: {str(e)}")

    # Test missing required property fields
    config['property2'] = {
        'address': 'test'  # Missing other required fields
    }
    try:
        validate_config(config)
        print("✗ Missing property fields test failed")
    except ValueError:
        print("✓ Missing property fields test passed")

if __name__ == "__main__":
    print("Running email validation tests...")
    test_email_validation()
    print("\nRunning config validation tests...")
    test_config_validation()