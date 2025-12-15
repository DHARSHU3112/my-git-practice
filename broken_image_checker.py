#!/usr/bin/env python3
"""
Broken Image Link Checker
This script crawls a website or checks a list of image URLs to identify broken images
and exports the results to an Excel file.
"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
import pandas as pd
from datetime import datetime
import argparse
import sys
from typing import List, Dict, Set
import time

class BrokenImageChecker:
    def __init__(self, timeout=10, max_retries=2):
        """
        Initialize the broken image checker.

        Args:
            timeout: Request timeout in seconds
            max_retries: Maximum number of retries for failed requests
        """
        self.timeout = timeout
        self.max_retries = max_retries
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.checked_images = set()
        self.results = []

    def check_image_url(self, image_url: str, source_page: str = "") -> Dict:
        """
        Check if an image URL is broken.

        Args:
            image_url: The image URL to check
            source_page: The page where the image was found

        Returns:
            Dictionary with check results
        """
        result = {
            'image_url': image_url,
            'source_page': source_page,
            'status': '',
            'status_code': '',
            'error_message': '',
            'checked_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        for attempt in range(self.max_retries + 1):
            try:
                response = self.session.head(
                    image_url,
                    timeout=self.timeout,
                    allow_redirects=True
                )

                # If HEAD request returns 405, try GET
                if response.status_code == 405:
                    response = self.session.get(
                        image_url,
                        timeout=self.timeout,
                        stream=True
                    )

                result['status_code'] = response.status_code

                if response.status_code == 200:
                    result['status'] = 'OK'
                elif response.status_code == 404:
                    result['status'] = 'BROKEN (Not Found)'
                elif response.status_code >= 400:
                    result['status'] = f'BROKEN (HTTP {response.status_code})'
                else:
                    result['status'] = f'WARNING (HTTP {response.status_code})'

                break

            except requests.exceptions.Timeout:
                result['status'] = 'BROKEN (Timeout)'
                result['error_message'] = 'Request timeout'
            except requests.exceptions.ConnectionError as e:
                result['status'] = 'BROKEN (Connection Error)'
                result['error_message'] = str(e)
            except requests.exceptions.TooManyRedirects:
                result['status'] = 'BROKEN (Too Many Redirects)'
                result['error_message'] = 'Too many redirects'
            except Exception as e:
                result['status'] = 'ERROR'
                result['error_message'] = str(e)

            # Wait before retry
            if attempt < self.max_retries:
                time.sleep(1)

        return result

    def extract_images_from_page(self, page_url: str) -> List[str]:
        """
        Extract all image URLs from a webpage.

        Args:
            page_url: The webpage URL to crawl

        Returns:
            List of image URLs found on the page
        """
        image_urls = []

        try:
            response = self.session.get(page_url, timeout=self.timeout)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'html.parser')

            # Find all img tags
            for img in soup.find_all('img'):
                img_src = img.get('src') or img.get('data-src')
                if img_src:
                    # Convert relative URLs to absolute
                    absolute_url = urljoin(page_url, img_src)
                    image_urls.append(absolute_url)

            # Also check for background images in CSS
            for element in soup.find_all(style=True):
                style = element.get('style', '')
                if 'background-image' in style or 'url(' in style:
                    # Simple extraction - can be enhanced
                    import re
                    urls = re.findall(r'url\([\'"]?([^\'"()]+)[\'"]?\)', style)
                    for url in urls:
                        absolute_url = urljoin(page_url, url)
                        if absolute_url.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.svg', '.webp')):
                            image_urls.append(absolute_url)

        except Exception as e:
            print(f"Error crawling {page_url}: {e}")

        return image_urls

    def check_images_from_url_list(self, image_urls: List[str], source_page: str = ""):
        """
        Check a list of image URLs.

        Args:
            image_urls: List of image URLs to check
            source_page: Optional source page reference
        """
        for image_url in image_urls:
            if image_url not in self.checked_images:
                print(f"Checking: {image_url}")
                result = self.check_image_url(image_url, source_page)
                self.results.append(result)
                self.checked_images.add(image_url)

                # Small delay to avoid overwhelming servers
                time.sleep(0.2)

    def check_images_from_webpage(self, page_url: str):
        """
        Extract and check all images from a webpage.

        Args:
            page_url: The webpage URL to check
        """
        print(f"\nExtracting images from: {page_url}")
        image_urls = self.extract_images_from_page(page_url)
        print(f"Found {len(image_urls)} images")

        self.check_images_from_url_list(image_urls, page_url)

    def export_to_excel(self, output_file: str):
        """
        Export results to an Excel file.

        Args:
            output_file: Path to the output Excel file
        """
        if not self.results:
            print("No results to export")
            return

        df = pd.DataFrame(self.results)

        # Create Excel writer with xlsxwriter engine
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Broken Images', index=False)

            # Get workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Broken Images']

            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })

            broken_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            })

            ok_format = workbook.add_format({
                'bg_color': '#C6EFCE',
                'font_color': '#006100'
            })

            warning_format = workbook.add_format({
                'bg_color': '#FFEB9C',
                'font_color': '#9C6500'
            })

            # Apply header format
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Set column widths
            worksheet.set_column('A:A', 60)  # image_url
            worksheet.set_column('B:B', 60)  # source_page
            worksheet.set_column('C:C', 25)  # status
            worksheet.set_column('D:D', 12)  # status_code
            worksheet.set_column('E:E', 40)  # error_message
            worksheet.set_column('F:F', 20)  # checked_at

            # Apply conditional formatting based on status
            for row_num, row_data in enumerate(self.results, start=1):
                status = row_data['status']
                if 'BROKEN' in status or 'ERROR' in status:
                    worksheet.set_row(row_num, None, broken_format)
                elif status == 'OK':
                    worksheet.set_row(row_num, None, ok_format)
                elif 'WARNING' in status:
                    worksheet.set_row(row_num, None, warning_format)

        print(f"\nResults exported to: {output_file}")

        # Print summary
        total = len(self.results)
        broken = sum(1 for r in self.results if 'BROKEN' in r['status'] or 'ERROR' in r['status'])
        ok = sum(1 for r in self.results if r['status'] == 'OK')

        print(f"\nSummary:")
        print(f"Total images checked: {total}")
        print(f"Broken images: {broken}")
        print(f"Working images: {ok}")
        print(f"Other (warnings): {total - broken - ok}")


def main():
    parser = argparse.ArgumentParser(
        description='Check for broken image links and export to Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Check images from a webpage
  python broken_image_checker.py --url https://example.com --output broken_images.xlsx

  # Check images from a text file (one URL per line)
  python broken_image_checker.py --file image_urls.txt --output broken_images.xlsx

  # Check a single image URL
  python broken_image_checker.py --image https://example.com/image.jpg --output result.xlsx
        """
    )

    parser.add_argument('--url', help='Website URL to crawl for images')
    parser.add_argument('--file', help='Text file containing image URLs (one per line)')
    parser.add_argument('--image', help='Single image URL to check')
    parser.add_argument('--output', '-o', default='broken_images.xlsx',
                       help='Output Excel file (default: broken_images.xlsx)')
    parser.add_argument('--timeout', type=int, default=10,
                       help='Request timeout in seconds (default: 10)')
    parser.add_argument('--retries', type=int, default=2,
                       help='Maximum retries for failed requests (default: 2)')

    args = parser.parse_args()

    # Validate input
    if not any([args.url, args.file, args.image]):
        parser.error('At least one of --url, --file, or --image is required')

    # Create checker instance
    checker = BrokenImageChecker(timeout=args.timeout, max_retries=args.retries)

    # Process based on input type
    if args.url:
        checker.check_images_from_webpage(args.url)

    if args.file:
        try:
            with open(args.file, 'r') as f:
                image_urls = [line.strip() for line in f if line.strip()]
            print(f"\nChecking {len(image_urls)} images from file: {args.file}")
            checker.check_images_from_url_list(image_urls, f"File: {args.file}")
        except FileNotFoundError:
            print(f"Error: File not found: {args.file}")
            sys.exit(1)

    if args.image:
        print(f"\nChecking single image: {args.image}")
        checker.check_images_from_url_list([args.image])

    # Export results
    checker.export_to_excel(args.output)


if __name__ == '__main__':
    main()
