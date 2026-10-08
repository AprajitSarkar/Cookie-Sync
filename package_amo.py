import os
import zipfile
import sys

def package_extension(source_dir, output_zip):
    source_dir = os.path.abspath(source_dir)
    output_zip = os.path.abspath(output_zip)
    
    if os.path.exists(output_zip):
        os.remove(output_zip)
        
    with zipfile.ZipFile(output_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, source_dir)
                # POSIX forward slash for AMO validation
                posix_arcname = rel_path.replace(os.path.sep, '/')
                
                # Create ZipInfo manually to guarantee forward slashes in header
                zinfo = zipfile.ZipInfo(posix_arcname)
                # Set standard Unix/POSIX compatible mode
                zinfo.create_system = 3 # Unix
                zinfo.external_attr = 0o644 << 16 # permissions
                zinfo.compress_type = zipfile.ZIP_DEFLATED
                
                with open(abs_path, 'rb') as f:
                    zf.writestr(zinfo, f.read())
                    
    print(f"Successfully created: {output_zip}")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    firefox_dir = os.path.join(base_dir, 'Firefox')
    chrome_dir = os.path.join(base_dir, 'Chrome')
    
    firefox_zip = os.path.join(base_dir, 'Cookie-Sync-Firefox.zip')
    chrome_zip = os.path.join(base_dir, 'Cookie-Sync-Chrome.zip')
    
    package_extension(firefox_dir, firefox_zip)
    package_extension(chrome_dir, chrome_zip)
