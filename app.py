from flask import Flask, render_template, request, jsonify, send_file
import requests
import re
import random
import csv
import io
from datetime import datetime, timedelta
from urllib.parse import quote

app = Flask(__name__)

# Common vendor names for random generation
COMMON_VENDORS = [
    'microsoft', 'google', 'apple', 'oracle', 'adobe', 'mozilla', 
    'cisco', 'ibm', 'intel', 'hp', 'dell', 'lenovo', 'asus',
    'samsung', 'sony', 'lg', 'vmware', 'redhat', 'canonical',
    'apache', 'nginx', 'nodejs', 'python', 'java', 'php'
]

PRODUCT_PREFIXES = ['server', 'client', 'pro', 'enterprise', 'professional', 'community', 'standard', 'ultimate']
PRODUCT_TYPES = ['suite', 'manager', 'viewer', 'editor', 'player', 'reader', 'browser', 'office', 'database', 'framework']

# CPE Dictionary - Sample CPE entries representing real-world software
CPE_DICTIONARY = [
    'cpe:2.3:a:microsoft:windows:10:*:*:*:*:*:*:*',
    'cpe:2.3:a:microsoft:windows:11:*:*:*:*:*:*:*',
    'cpe:2.3:a:microsoft:office:2019:*:*:*:*:*:*:*',
    'cpe:2.3:a:microsoft:office:2021:sp1:*:*:professional:*:*:*',
    'cpe:2.3:a:microsoft:edge:120.0.2210.91:*:*:*:*:*:*:*',
    'cpe:2.3:a:google:chrome:120.0.6099.129:*:*:*:*:*:*:*',
    'cpe:2.3:a:google:chrome:119.0.6045.199:stable:*:*:*:*:*:*',
    'cpe:2.3:a:mozilla:firefox:121.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:mozilla:firefox:120.0.1:*:*:*:*:linux:*:*',
    'cpe:2.3:a:mozilla:thunderbird:115.6.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:adobe:acrobat_reader:23.008.20458:*:*:*:*:*:*:*',
    'cpe:2.3:a:adobe:photoshop:24.7.1:*:*:*:*:*:*:*',
    'cpe:2.3:a:adobe:illustrator:28.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:oracle:java:1.8.0:update391:*:*:*:*:*:*',
    'cpe:2.3:a:oracle:java:17.0.9:*:*:*:*:*:*:*',
    'cpe:2.3:a:oracle:mysql:8.0.35:*:*:*:*:*:*:*',
    'cpe:2.3:a:oracle:virtualbox:7.0.12:*:*:*:*:*:*:*',
    'cpe:2.3:a:apache:tomcat:10.1.17:*:*:*:*:*:*:*',
    'cpe:2.3:a:apache:http_server:2.4.58:*:*:*:*:*:*:*',
    'cpe:2.3:a:apache:maven:3.9.6:*:*:*:*:*:*:*',
    'cpe:2.3:a:python:python:3.12.1:*:*:*:*:*:*:*',
    'cpe:2.3:a:python:python:3.11.7:*:*:*:*:*:*:*',
    'cpe:2.3:a:nodejs:node.js:20.10.0:*:*:*:lts:*:*:*',
    'cpe:2.3:a:nodejs:node.js:21.5.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:php:php:8.3.1:*:*:*:*:*:*:*',
    'cpe:2.3:a:php:php:8.2.14:*:*:*:*:*:*:*',
    'cpe:2.3:a:vmware:workstation:17.5.0:*:*:*:pro:*:*:*',
    'cpe:2.3:a:vmware:vsphere:8.0:update2:*:*:*:*:*:*',
    'cpe:2.3:a:cisco:webex:43.12.0.27982:*:*:*:*:*:*:*',
    'cpe:2.3:a:cisco:anyconnect:4.10.07061:*:*:*:*:*:*:*',
    'cpe:2.3:a:ibm:db2:11.5.8:*:*:*:*:linux:*:*',
    'cpe:2.3:a:ibm:websphere:9.0.5.15:*:*:*:*:*:*:*',
    'cpe:2.3:a:intel:graphics_driver:31.0.101.4972:*:*:*:*:windows:*:*',
    'cpe:2.3:a:nvidia:geforce_experience:3.27.0.112:*:*:*:*:*:*:*',
    'cpe:2.3:a:zoom:zoom:5.16.10.26186:*:*:*:*:windows:*:*',
    'cpe:2.3:a:slack:slack:4.36.140:*:*:*:*:*:*:*',
    'cpe:2.3:a:docker:docker:24.0.7:*:*:*:*:*:*:*',
    'cpe:2.3:a:git:git:2.43.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:7-zip:7-zip:23.01:*:*:*:*:*:*:*',
    'cpe:2.3:a:videolan:vlc:3.0.20:*:*:*:*:*:*:*',
    'cpe:2.3:a:wireshark:wireshark:4.2.0:*:*:*:*:*:*:*',
    'cpe:2.3:a:notepad_plus_plus:notepad\\+\\+:8.6.2:*:*:*:*:*:*:*',
    'cpe:2.3:a:postman:postman:10.21.1:*:*:*:*:*:*:*',
    'cpe:2.3:a:jetbrains:intellij_idea:2023.3.2:*:*:*:community:*:*:*',
    'cpe:2.3:a:visual_studio_code:visual_studio_code:1.85.2:*:*:*:*:*:*:*',
    'cpe:2.3:a:spotify:spotify:1.2.26.1187:*:*:*:*:*:*:*',
    'cpe:2.3:a:steam:steam:1702689516:*:*:*:*:*:*:*',
    'cpe:2.3:a:discord:discord:0.0.309:*:*:*:*:*:*:*',
    'cpe:2.3:a:malwarebytes:anti-malware:4.6.3:*:*:*:premium:*:*:*',
    'cpe:2.3:a:ccleaner:ccleaner:6.19.10858:*:*:*:free:*:*:*'
]

def parse_cpe_uri(cpe_string):
    """
    Parse CPE URI format (cpe:2.3:a:vendor:product:version:...)
    Returns dict with parsed components
    """
    try:
        if not cpe_string.startswith('cpe:'):
            return None
        
        parts = cpe_string.split(':')
        if len(parts) < 5:
            return None
        
        result = {
            'cpe': cpe_string,
            'vendor': parts[3] if len(parts) > 3 and parts[3] != '*' else '',
            'product': parts[4] if len(parts) > 4 and parts[4] != '*' else '',
            'version': parts[5] if len(parts) > 5 and parts[5] != '*' else '',
            'update': parts[6] if len(parts) > 6 and parts[6] != '*' else '',
            'edition': parts[7] if len(parts) > 7 and parts[7] != '*' else '',
            'language': parts[8] if len(parts) > 8 and parts[8] != '*' else '',
            'sw_edition': parts[9] if len(parts) > 9 and parts[9] != '*' else '',
            'target_sw': parts[10] if len(parts) > 10 and parts[10] != '*' else '',
            'target_hw': parts[11] if len(parts) > 11 and parts[11] != '*' else '',
        }
        
        # Create other fields description
        other_fields = []
        if result['update']:
            other_fields.append(f"Update: {result['update']}")
        if result['edition']:
            other_fields.append(f"Edition: {result['edition']}")
        if result['language']:
            other_fields.append(f"Language: {result['language']}")
        if result['sw_edition']:
            other_fields.append(f"SW Edition: {result['sw_edition']}")
        if result['target_sw']:
            other_fields.append(f"Target SW: {result['target_sw']}")
        if result['target_hw']:
            other_fields.append(f"Target HW: {result['target_hw']}")
        
        result['other_fields'] = ', '.join(other_fields) if other_fields else 'None'
        
        return result
    except Exception as e:
        print(f"Error parsing CPE: {e}")
        return None

def generate_installation_metadata():
    """Generate simulated installation metadata"""
    # Random size between 10MB and 2000MB
    size_mb = round(random.uniform(10, 2000), 2)
    
    # Random installation date within last 2 years
    days_ago = random.randint(0, 730)
    install_date = (datetime.now() - timedelta(days=days_ago)).strftime('%Y-%m-%d')
    
    # Installation location (default C drive)
    locations = [
        'C:\\Program Files\\',
        'C:\\Program Files (x86)\\',
        'C:\\Users\\Public\\',
        'C:\\ProgramData\\'
    ]
    location = random.choice(locations)
    
    return {
        'size_mb': size_mb,
        'install_date': install_date,
        'install_location': location
    }

def validate_cpe_with_nvd(cpe_string):
    """
    Validate CPE against NVD CPE dictionary
    Returns True if valid, False otherwise
    """
    try:
        # Try to search for the CPE in NVD
        # Note: In production, you would use NVD API with proper rate limiting
        # For this implementation, we'll do basic validation
        
        # Basic format validation
        if not cpe_string.startswith('cpe:2.3:'):
            return False
        
        parts = cpe_string.split(':')
        if len(parts) < 6:
            return False
        
        # Check if vendor and product are not empty
        if parts[3] == '*' or parts[4] == '*':
            return False
        
        return True
    except Exception as e:
        print(f"Error validating CPE: {e}")
        return False

def search_nvd_cpe(vendor, product, version=''):
    """
    Search NVD for CPE entries matching vendor, product, and optionally version
    Returns list of matching CPEs
    """
    try:
        # Note: This is a simplified implementation
        # In production, use the official NVD CPE API: https://nvd.nist.gov/developers/products
        
        # For demonstration, generate some sample CPEs
        cpes = []
        
        # Generate a few variations
        base_cpe = f"cpe:2.3:a:{vendor}:{product}:{version if version else '*'}"
        for i in range(3):
            v = version if version else f"{random.randint(1,10)}.{random.randint(0,9)}.{random.randint(0,9)}"
            cpe = f"cpe:2.3:a:{vendor}:{product}:{v}:*:*:*:*:*:*:*"
            cpes.append(cpe)
        
        return cpes
    except Exception as e:
        print(f"Error searching NVD: {e}")
        return []

def generate_random_cpe():
    """Generate random CPE with metadata"""
    vendor = random.choice(COMMON_VENDORS)
    
    # Generate product name
    if random.random() > 0.5:
        product = f"{random.choice(PRODUCT_PREFIXES)}_{random.choice(PRODUCT_TYPES)}"
    else:
        product = f"{vendor}_{random.choice(PRODUCT_TYPES)}"
    
    # Generate version
    major = random.randint(1, 20)
    minor = random.randint(0, 9)
    patch = random.randint(0, 99)
    version = f"{major}.{minor}.{patch}"
    
    # Create CPE string
    cpe_string = f"cpe:2.3:a:{vendor}:{product}:{version}:*:*:*:*:*:*:*"
    
    # Check if similar CPE exists and adjust if needed
    similar_cpes = search_nvd_cpe(vendor, product)
    if similar_cpes and random.random() > 0.5:
        # Sometimes use a similar existing CPE
        cpe_string = random.choice(similar_cpes)
    
    return {
        'vendor': vendor,
        'product': product,
        'version': version,
        'cpe': cpe_string,
        **generate_installation_metadata()
    }

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/api/auto-fetch-cpe', methods=['POST'])
def auto_fetch_cpe():
    """
    Auto-fetch CPE entries from CPE dictionary
    Expected input: count (number of CPE entries to fetch, default 10)
    """
    try:
        data = request.json
        count = data.get('count', 10)
        count = min(max(1, count), 100)  # Limit between 1 and 100
        
        # Randomly select CPE entries from the dictionary
        selected_cpes = random.sample(CPE_DICTIONARY, min(count, len(CPE_DICTIONARY)))
        
        results = []
        for cpe_string in selected_cpes:
            # Validate CPE
            if validate_cpe_with_nvd(cpe_string):
                # Parse CPE
                parsed = parse_cpe_uri(cpe_string)
                if parsed:
                    # Add installation metadata
                    metadata = generate_installation_metadata()
                    results.append({**parsed, **metadata})
        
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/fetch-cpe', methods=['POST'])
def fetch_cpe():
    """
    Fetch and parse CPE from user input
    Expected input: CPE string or search parameters
    """
    try:
        data = request.json
        cpe_input = data.get('cpe_string', '')
        
        if not cpe_input:
            return jsonify({'error': 'CPE string is required'}), 400
        
        # Validate CPE
        if not validate_cpe_with_nvd(cpe_input):
            return jsonify({'error': 'Invalid CPE format or CPE not found in dictionary'}), 400
        
        # Parse CPE
        parsed = parse_cpe_uri(cpe_input)
        if not parsed:
            return jsonify({'error': 'Failed to parse CPE'}), 400
        
        # Add installation metadata
        metadata = generate_installation_metadata()
        result = {**parsed, **metadata}
        
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/search-cpe', methods=['POST'])
def search_cpe():
    """
    Search for CPE entries based on vendor/product
    """
    try:
        data = request.json
        vendor = data.get('vendor', '')
        product = data.get('product', '')
        
        if not vendor or not product:
            return jsonify({'error': 'Vendor and product are required'}), 400
        
        # Search for CPEs
        cpes = search_nvd_cpe(vendor, product)
        
        results = []
        for cpe in cpes[:10]:  # Limit to 10 results
            parsed = parse_cpe_uri(cpe)
            if parsed:
                metadata = generate_installation_metadata()
                results.append({**parsed, **metadata})
        
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-random', methods=['POST'])
def generate_random():
    """
    Generate random CPE entries
    """
    try:
        data = request.json
        count = data.get('count', 5)
        count = min(count, 50)  # Limit to 50 entries
        
        results = []
        for _ in range(count):
            cpe_data = generate_random_cpe()
            results.append(cpe_data)
        
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export-csv', methods=['POST'])
def export_csv():
    """
    Export CPE data to CSV
    """
    try:
        data = request.json
        cpe_data = data.get('data', [])
        
        if not cpe_data:
            return jsonify({'error': 'No data to export'}), 400
        
        # Create CSV in memory
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write header
        writer.writerow([
            'CPE', 'Vendor', 'Product', 'Version', 
            'Other Fields', 'Size (MB)', 'Install Date', 'Install Location'
        ])
        
        # Write data
        for item in cpe_data:
            writer.writerow([
                item.get('cpe', ''),
                item.get('vendor', ''),
                item.get('product', ''),
                item.get('version', ''),
                item.get('other_fields', ''),
                item.get('size_mb', ''),
                item.get('install_date', ''),
                item.get('install_location', '')
            ])
        
        # Prepare response
        output.seek(0)
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f'cpe_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    # Note: For production deployment, set debug=False and use a production WSGI server
    import os
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    app.run(debug=debug_mode, host='0.0.0.0', port=5000)
