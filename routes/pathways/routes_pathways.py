from flask import Blueprint, jsonify
from flask_login import login_required
from simple_salesforce import Salesforce, SalesforceAuthenticationFailed
from models import db
from models.pathways import Pathway
from config import Config

pathways_bp = Blueprint('pathways', __name__, url_prefix='/pathways')

@pathways_bp.route('/import-from-salesforce', methods=['POST'])
@login_required
def import_pathways_from_salesforce():
    try:
        print("Fetching pathway data from Salesforce...")
        success_count = 0
        error_count = 0
        errors = []

        # Connect to Salesforce
        sf = Salesforce(
            username=Config.SF_USERNAME,
            password=Config.SF_PASSWORD,
            security_token=Config.SF_SECURITY_TOKEN,
            domain='login'
        )

        # Query pathways from Salesforce
        pathways_query = """
        SELECT Id, Name
        FROM Pathway__c
        """

        # Execute query
        pathways_result = sf.query_all(pathways_query)
        pathway_rows = pathways_result.get('records', [])

        # Process pathways
        for row in pathway_rows:
            try:
                # Check if pathway already exists
                pathway = Pathway.query.filter_by(salesforce_id=row['Id']).first()
                
                if pathway:
                    # Update existing pathway
                    pathway.name = row['Name']
                else:
                    # Create new pathway
                    pathway = Pathway(
                        salesforce_id=row['Id'],
                        name=row['Name']
                    )
                    db.session.add(pathway)
                
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_msg = f"Error processing pathway {row.get('Name', 'Unknown')}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)

        # Commit all changes
        db.session.commit()
        
        # Print summary
        print(f"\nSuccessfully processed {success_count} pathways with {error_count} errors")
        if errors:
            print("\nErrors encountered:")
            for error in errors:
                print(f"- {error}")
        
        return jsonify({
            'success': True,
            'message': f'Successfully processed {success_count} pathways with {error_count} errors',
            'errors': errors
        })

    except SalesforceAuthenticationFailed:
        print("Error: Failed to authenticate with Salesforce")
        return jsonify({
            'success': False,
            'message': 'Failed to authenticate with Salesforce'
        }), 401
    except Exception as e:
        db.session.rollback()
        print(f"Error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

