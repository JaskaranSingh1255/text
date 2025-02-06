import json
import pandas as pd
from openai import OpenAI  # Import OpenAI library

# Initialize OpenAI client
client = OpenAI(api_key=""
                )  # Replace with your actual API key

# Function to read responses from Excel and process them
def read_responses_from_excel(file_path, sheet_name='Q6_Industry', column_name='UserRes'):
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Extract the responses from the specified column
    user_responses = df[column_name].dropna().tolist()
    
    return user_responses

def process_tickets(survey_question, ticket_categories, user_responses, output_file="Industry.xlsx"):
    results = []
    i = 0
    for user_response in user_responses:
        prompt = f"""
        TASK: Categorize the user response into predefined categories, subcategories, and sub-subcategories.

        PARAMETERS:
        1. **Survey Question**: {survey_question}
        2. **Ticket Categories**: {json.dumps(ticket_categories, indent=2)}
        3. **User Response**: {user_response}

        **Output Format** (Return as JSON):
        {{
            "Category": "Main category",
            "Subcategory": "Subcategory",
            "Sub-Subcategory": "Sub-subcategory"
        }}
        """

        messages = [
            {"role": "system", "content": "You are a helpful assistant that classifies user responses."},
            {"role": "user", "content": prompt},
        ]

        # API call with correct response format
        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            response_format={"type": "json_object"}  # ✅ Corrected value
        )

        # Extract structured JSON response
        parsed_response = completion.choices[0].message.content
        print("Raw JSON String:", parsed_response)
        
        # Convert JSON string to Python dictionary
        try:
            parsed_response = json.loads(parsed_response)
        except json.JSONDecodeError as e:
            print("JSON Decode Error:", e)

        # Append results
        results.append({
            "User Response": user_response,
            "Category": parsed_response.get("Category", "Uncategorized"),
            "Subcategory": parsed_response.get("Subcategory", "Uncategorized"),
            "Sub-Subcategory": parsed_response.get("Sub-Subcategory", "Uncategorized"),
        })
        print("response", str(i))
        i += 1

    # Save results to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# Example usage
survey_question = "Describe your issue briefly."
ticket_categories = {
    "Positive Feedback": {
        "Product & Order Experience": {
            "Quality Products": {
                "Accurate product descriptions and high quality": "Subcategory 1",
                "Durable and well-packaged items": "Subcategory 1"
            },
            "Wide Variety": {
                "Good selection of brands and categories": "Subcategory 2",
                "Easy comparison between similar products": "Subcategory 2"
            },
            "On-Time Delivery": {
                "Orders delivered before or on expected date": "Subcategory 3",
                "Well-tracked and properly packaged deliveries": "Subcategory 3"
            }
        },
        "Customer Service & Returns": {
            "Helpful Support": {
                "Responsive and knowledgeable customer support": "Subcategory 1",
                "Quick resolution of issues and queries": "Subcategory 1"
            },
            "Easy Returns & Refunds": {
                "Hassle-free return process": "Subcategory 2",
                "Fast refund and replacement policy": "Subcategory 2"
            },
            "Loyalty & Discounts": {
                "Good reward points and membership perks": "Subcategory 3",
                "Attractive discounts and seasonal offers": "Subcategory 3"
            }
        },
        "App & Payment Experience": {
            "User-Friendly App": {
                "Smooth navigation and checkout process": "Subcategory 1",
                "Good filter and search functionality": "Subcategory 1"
            },
            "Multiple Payment Options": {
                "Support for credit cards, UPI, COD, etc.": "Subcategory 2",
                "Secure and fast transactions": "Subcategory 2"
            },
            "Tracking & Notifications": {
                "Real-time order tracking updates": "Subcategory 3",
                "Helpful delivery status alerts": "Subcategory 3"
            }
        }
    },
    "Negative Feedback": {
        "Product & Order Issues": {
            "Defective or Wrong Products": {
                "Received damaged or defective item": "Subcategory 1",
                "Incorrect product or size sent": "Subcategory 1"
            },
            "Poor Packaging": {
                "Items received in bad condition": "Subcategory 2",
                "Spillage or breakage due to weak packaging": "Subcategory 2"
            },
            "Late or Missing Orders": {
                "Delayed shipment beyond expected delivery date": "Subcategory 3",
                "Order lost or missing in transit": "Subcategory 3"
            }
        },
        "Customer Service & Refund Issues": {
            "Unresponsive Support": {
                "Delayed response from customer care": "Subcategory 1",
                "Unsatisfactory issue resolution": "Subcategory 1"
            },
            "Return/Refund Problems": {
                "Complicated or rejected return requests": "Subcategory 2",
                "Delayed refund processing": "Subcategory 2"
            },
            "Discount & Coupon Issues": {
                "Coupons not applied properly": "Subcategory 3",
                "Misleading discounts or hidden charges": "Subcategory 3"
            }
        },
        "App & Payment Issues": {
            "App Errors & Bugs": {
                "Crashes or slow app performance": "Subcategory 1",
                "Login/logout problems": "Subcategory 1"
            },
            "Payment Failures": {
                "Transaction failures and incorrect deductions": "Subcategory 2",
                "Delayed refund for failed transactions": "Subcategory 2"
            },
            "Tracking & Notification Issues": {
                "Incorrect tracking updates": "Subcategory 3",
                "No proper alerts for order status": "Subcategory 3"
            }
        }
    }
}






# Replace 'your_file.xlsx' with the path to your actual Excel file}
file_path = '15.Onlinepurchases.xlsx'
user_responses = read_responses_from_excel(file_path)

process_tickets(survey_question, ticket_categories, user_responses)
