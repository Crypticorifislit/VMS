from flask import Blueprint, render_template, request, flash, jsonify, make_response, send_file
from flask_login import login_required
from sqlalchemy import any_, extract
from models.event import Event, EventAttendance, EventType, EventStatus
from datetime import datetime, timedelta
from models.school_model import School
from models.teacher import Teacher
from models.upcoming_events import UpcomingEvent
from models.volunteer import Volunteer, EventParticipation, Skill, VolunteerSkill
from models.organization import Organization, VolunteerOrganization
from models.event import event_districts
from models.district_model import District
from models.reports import DistrictYearEndReport
from models import db  # Import db from models instead of creating new instance
from models.pathways import Pathway
import json
import pytz
import pandas as pd
import io
import xlsxwriter

report_bp = Blueprint('report', __name__)

# District mapping configuration
DISTRICT_MAPPING = {
    # Districts to show
    '0015f00000JU4opAAD': {  # Grandview School District
        'name': 'Grandview School District',
        'district_code': '48074',
        'merge_with': None,
        'show': True,
        'aliases': [
            'Grandview',
        ]
    },
    '0015f00000JU4pVAAT': {  # Kansas City Kansas Public Schools
        'name': 'Kansas City Kansas Public Schools',
        'district_code': '48078',
        'merge_with': None,
        'show': True,
        'aliases': [
            'Kansas City Kansas School District',
            'KANSAS CITY USD 500',
            'KCKPS (KS)'
        ]
    },
    '0015f00000JU9ZEAA1': {  # Allen Village - District
        'name': 'Allen Village - District',
        'district_code': '48909',
        'merge_with': None,
        'show': True
    },
    '0015f00000JU9ZFAA1': {  # Hickman Mills School District
        'name': 'Hickman Mills School District',
        'district_code': '48072',
        'merge_with': None,
        'show': True,
        'aliases': ['Hickman Mills']
    },
    '0015f00000JVaPKAA1': {  # GUADALUPE CENTERS SCHOOLS
        'name': 'GUADALUPE CENTERS SCHOOLS',
        'district_code': '48902',
        'merge_with': None,
        'show': True
    },
    '0015f00000KvuZTAAZ': {  # Center School District
        'name': 'Center School District',
        'district_code': '48080',
        'merge_with': None,
        'show': True
    },
    '0015f00000KxHwVAAV': {  # Kansas City Public Schools (MO)
        'name': 'Kansas City Public Schools (MO)',
        'district_code': None,
        'merge_with': None,
        'show': True,
        'aliases': ['KCPS (MO)']
    }
}

@report_bp.route('/reports')
@login_required
def reports():
    # Define available reports
    available_reports = [
        {
            'title': 'Virtual Session Usage',
            'description': 'View virtual session statistics by district, including attendance rates and total participation.',
            'icon': 'fa-solid fa-video',
            'url': '/reports/virtual/usage',
            'category': 'Virtual Events'
        },
        {
            'title': 'Volunteer Thank You Report',
            'description': 'View top volunteers by hours and events for end of year thank you notes.',
            'icon': 'fa-solid fa-heart',
            'url': '/reports/volunteer/thankyou',
            'category': 'Volunteer Recognition'
        },
        {
            'title': 'Organization Thank You Report',
            'description': 'View organization contributions by sessions, hours, and volunteer participation.',
            'icon': 'fa-solid fa-building',
            'url': '/reports/organization/thankyou',
            'category': 'Organization Recognition'
        },
        {
            'title': 'District Year-End Report',
            'description': 'View comprehensive year-end statistics for each district.',
            'icon': 'fa-solid fa-chart-pie',
            'url': '/reports/district/year-end',
            'category': 'District Reports'
        },
        {
            'title': 'Recruitment Tools',
            'description': 'Access various tools for volunteer recruitment and event matching.',
            'icon': 'fa-solid fa-users',
            'url': '/reports/recruitment',
            'category': 'Recruitment'
        },
        {
            'title': 'Event Contact Report',
            'description': 'View upcoming events and volunteer contact information.',
            'icon': 'fa-solid fa-address-book',
            'url': '/reports/contact',
            'category': 'Event Management'
        },
        {
            'title': 'Pathway Report',
            'description': 'View pathway statistics including student participation, events, and contact engagement.',
            'icon': 'fa-solid fa-road',
            'url': '/reports/pathways',
            'category': 'Program Reports'
        },
        {
            'title': 'New Skills Summary Report',
            'description': 'View new skills learned by students in events.',
            'icon': 'fa-solid fa-graduation-cap',
            'url': '/reports/NewSkillsReports.html',
            'category': 'Student Learning'
        }
    ]
    
    return render_template('reports/reports.html', reports=available_reports)


@report_bp.route('/reports/NewSkillsReports.html')
@login_required
def Skillreport():
    # Get the 'name' filter from the query string (if any)
    name_filter = request.args.get('name', None)

    # Get pagination parameters
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 25, type=int)
    
    # Validate per_page value
    allowed_per_page = [10, 25, 50, 100]
    if per_page not in allowed_per_page:
        per_page = 25  # fallback to default if invalid value
   
    query = Volunteer.query
    name = None
    query = query.filter(Volunteer.organization_name.is_((None)))
    # Apply pagination
    pagination = query.paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )
    
    # Paginate the query (x per page)
    
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    
    #If the name_filter is provided, filter the query
    current_filters = {
       'organization_name': Organization.name
    }

    # Get the items for the current page
    volunteers = pagination.items
    print(name_filter)
    return render_template(
        'reports/NewSkillsReports.html',
        volunteers=volunteers,
        pagination=pagination,
        name_filter=name_filter,  # Pass this so you can keep it in the search box or page links
        current_filters= current_filters 
    )

@report_bp.route('/reports/virtual/usage')
@login_required
def virtual_usage():
    # Get filter parameters with explicit year handling
    year = int(request.args.get('year', '2024'))  # Default to 2024, convert to int
    
    # Check if date_from and date_to are provided and match the selected year
    date_from_str = request.args.get('date_from')
    date_to_str = request.args.get('date_to')
    
    # If dates are provided, check if they match the selected year
    if date_from_str and date_to_str:
        try:
            date_from = datetime.strptime(date_from_str, '%Y-%m-%d')
            date_to = datetime.strptime(date_to_str, '%Y-%m-%d')
            # If the dates don't match the selected year, use the year's full range
            if date_from.year != year or date_to.year != year:
                date_from = datetime(year, 1, 1)
                date_to = datetime(year, 12, 31)
        except ValueError:
            date_from = datetime(year, 1, 1)
            date_to = datetime(year, 12, 31)
    else:
        # If no dates provided, use the full year range
        date_from = datetime(year, 1, 1)
        date_to = datetime(year, 12, 31)

    current_filters = {
        'year': year,
        'date_from': date_from,
        'date_to': date_to,
        'career_cluster': request.args.get('career_cluster'),
        'school': request.args.get('school'),
        'district': request.args.get('district')
    }
    
    # Base query with date range filter
    query = Event.query.filter(
        Event.type == EventType.VIRTUAL_SESSION,
        Event.status == EventStatus.COMPLETED,
        Event.start_date >= date_from,
        Event.start_date <= date_to
    )
    
    # Apply additional filters
    if current_filters['career_cluster']:
        query = query.filter(Event.series.ilike(f'%{current_filters["career_cluster"]}%'))
    if current_filters['school']:
        query = query.filter(Event.school.ilike(f'%{current_filters["school"]}%'))
    if current_filters['district']:
        query = query.filter(Event.district_partner.ilike(f'%{current_filters["district"]}%'))
    
    # Get all events after filtering
    virtual_events = query.all()
    
    # Get unique values for filter dropdowns
    all_clusters = db.session.query(Event.series).distinct().all()
    all_schools = db.session.query(Event.school).distinct().all()
    all_districts = db.session.query(Event.district_partner).distinct().all()
    
    filter_options = {
        'career_clusters': sorted([c[0] for c in all_clusters if c[0]]),
        'schools': sorted([s[0] for s in all_schools if s[0]]),
        'districts': sorted([d[0] for d in all_districts if d[0]])
    }
    
    # Create a dictionary to store district stats
    district_stats = {}
    
    for event in virtual_events:
        district_name = event.district_partner or 'Unassigned'
        
        if district_name not in district_stats:
            district_stats[district_name] = {
                'total_sessions': 0,
                'total_registered': 0,
                'total_attended': 0,
                'avg_attendance_rate': 0,
                'total_duration': 0,
                'schools': set(),
                'educators': set(),
                'career_clusters': set()
            }
        
        stats = district_stats[district_name]
        stats['total_sessions'] += 1
        stats['total_registered'] += event.registered_count or 0
        stats['total_attended'] += event.attended_count or 0
        stats['total_duration'] += event.duration or 0
        
        if event.school:
            stats['schools'].add(event.school)
        if event.educators:
            stats['educators'].add(event.educators)
        if event.series:
            stats['career_clusters'].add(event.series)
    
    # Calculate averages and convert sets to counts
    for stats in district_stats.values():
        if stats['total_registered'] > 0:
            stats['avg_attendance_rate'] = (stats['total_attended'] / stats['total_registered']) * 100
        stats['avg_duration'] = stats['total_duration'] / stats['total_sessions'] if stats['total_sessions'] > 0 else 0
        stats['school_count'] = len(stats['schools'])
        stats['educator_count'] = len(stats['educators'])
        stats['career_cluster_count'] = len(stats['career_clusters'])
        
        # Clean up sets before sending to template
        del stats['schools']
        del stats['educators']
        del stats['career_clusters']
    
    # Sort districts by total sessions
    sorted_districts = dict(sorted(
        district_stats.items(), 
        key=lambda x: x[1]['total_sessions'], 
        reverse=True
    ))
    
    return render_template(
        'reports/virtual_usage.html',
        district_stats=sorted_districts,
        filter_options=filter_options,
        current_filters=current_filters
    )

@report_bp.route('/reports/virtual/usage/district/<district_name>')
@login_required
def virtual_usage_district(district_name):
    # Get filter parameters with explicit year handling
    year = int(request.args.get('year', '2024'))  # Default to 2024, convert to int
    
    # Check if date_from and date_to are provided and match the selected year
    date_from_str = request.args.get('date_from')
    date_to_str = request.args.get('date_to')
    
    # If dates are provided, check if they match the selected year
    if date_from_str and date_to_str:
        try:
            date_from = datetime.strptime(date_from_str, '%Y-%m-%d')
            date_to = datetime.strptime(date_to_str, '%Y-%m-%d')
            # If the dates don't match the selected year, use the year's full range
            if date_from.year != year or date_to.year != year:
                date_from = datetime(year, 1, 1)
                date_to = datetime(year, 12, 31)
        except ValueError:
            date_from = datetime(year, 1, 1)
            date_to = datetime(year, 12, 31)
    else:
        # If no dates provided, use the full year range
        date_from = datetime(year, 1, 1)
        date_to = datetime(year, 12, 31)

    current_filters = {
        'year': year,
        'date_from': date_from,
        'date_to': date_to
    }
    
    # Base query for this district with date filtering
    query = Event.query.filter(
        Event.type == EventType.VIRTUAL_SESSION,
        Event.status == EventStatus.COMPLETED,
        Event.district_partner == district_name,
        Event.start_date >= date_from,
        Event.start_date <= date_to
    ).order_by(Event.start_date)
    
    # Get all events for this district
    district_events = query.all()
    
    # Group events by month
    monthly_stats = {}
    
    for event in district_events:
        month_key = event.start_date.strftime('%Y-%m')
        
        if month_key not in monthly_stats:
            monthly_stats[month_key] = {
                'month_name': event.start_date.strftime('%B %Y'),
                'total_sessions': 0,
                'total_registered': 0,
                'total_attended': 0,
                'total_duration': 0,
                'schools': set(),
                'educators': set(),
                'career_clusters': set(),
                'events': []
            }
        
        stats = monthly_stats[month_key]
        stats['total_sessions'] += 1
        stats['total_registered'] += event.registered_count or 0
        stats['total_attended'] += event.attended_count or 0
        stats['total_duration'] += event.duration or 0
        
        if event.school:
            stats['schools'].add(event.school)
        if event.educators:
            stats['educators'].add(event.educators)
        if event.series:
            stats['career_clusters'].add(event.series)
            
        # Add event details
        stats['events'].append({
            'title': event.title,
            'date': event.start_date.strftime('%Y-%m-%d'),
            'duration': event.duration,
            'registered': event.registered_count,
            'attended': event.attended_count,
            'school': event.school,
            'educator': event.educators,
            'career_cluster': event.series
        })
    
    # Calculate averages and convert sets to counts
    for stats in monthly_stats.values():
        if stats['total_registered'] > 0:
            stats['avg_attendance_rate'] = (stats['total_attended'] / stats['total_registered']) * 100
        stats['avg_duration'] = stats['total_duration'] / stats['total_sessions'] if stats['total_sessions'] > 0 else 0
        stats['school_count'] = len(stats['schools'])
        stats['educator_count'] = len(stats['educators'])
        stats['career_cluster_count'] = len(stats['career_clusters'])
        
        # Clean up sets
        del stats['schools']
        del stats['educators']
        del stats['career_clusters']
    
    # Sort months chronologically
    sorted_stats = dict(sorted(monthly_stats.items()))
    
    return render_template(
        'reports/virtual_usage_district.html',
        district_name=district_name,
        monthly_stats=sorted_stats,
        current_filters=current_filters
    )

@report_bp.route('/reports/volunteer/thankyou')
@login_required
def volunteer_thankyou():
    # Get filter parameters
    year = request.args.get('year', datetime.now().year)
    
    # Query volunteer participation through EventParticipation
    volunteer_stats = db.session.query(
        Volunteer,
        db.func.sum(EventParticipation.delivery_hours).label('total_hours'),
        db.func.count(EventParticipation.id).label('total_events'),
        Organization.name.label('organization_name')
    ).join(
        EventParticipation, Volunteer.id == EventParticipation.volunteer_id
    ).join(
        Event, EventParticipation.event_id == Event.id
    ).outerjoin(
        VolunteerOrganization, Volunteer.id == VolunteerOrganization.volunteer_id
    ).outerjoin(
        Organization, VolunteerOrganization.organization_id == Organization.id
    ).filter(
        extract('year', Event.start_date) == year,
        EventParticipation.status == 'Attended'
    ).group_by(
        Volunteer.id,
        Organization.name
    ).order_by(
        db.desc('total_hours')
    ).all()

    # Format the data for the template
    volunteer_data = [{
        'id': v.id,
        'name': f"{v.first_name} {v.last_name}",
        'total_hours': round(float(hours or 0), 2) if hours is not None else 0,
        'total_events': events,
        'organization': org or 'Independent'
    } for v, hours, events, org in volunteer_stats]

    return render_template(
        'reports/volunteer_thankyou.html',
        volunteers=volunteer_data,
        year=year,
        now=datetime.now()
    )

@report_bp.route('/reports/volunteer/thankyou/detail/<int:volunteer_id>')
@login_required
def volunteer_thankyou_detail(volunteer_id):
    # Get filter parameters
    year = request.args.get('year', datetime.now().year)
    
    # Get volunteer details
    volunteer = Volunteer.query.get_or_404(volunteer_id)
    
    # Query all events for this volunteer in the specified year
    events = db.session.query(
        Event,
        EventParticipation.delivery_hours
    ).join(
        EventParticipation, Event.id == EventParticipation.event_id
    ).filter(
        EventParticipation.volunteer_id == volunteer_id,
        extract('year', Event.start_date) == year,
        EventParticipation.status == 'Attended'
    ).order_by(
        Event.start_date
    ).all()
    
    # Format the events data
    events_data = [{
        'date': event.start_date.strftime('%B %d, %Y'),
        'title': event.title,
        'type': event.type.value if event.type else 'Unknown',
        'hours': round(hours or 0, 2),
        'school': event.school or 'N/A',
        'district': event.district_partner or 'N/A'
    } for event, hours in events]
    
    # Calculate totals
    total_hours = sum(event['hours'] for event in events_data)
    total_events = len(events_data)
    
    return render_template(
        'reports/volunteer_thankyou_detail.html',
        volunteer=volunteer,
        events=events_data,
        total_hours=total_hours,
        total_events=total_events,
        year=year,
        now=datetime.now()
    )

@report_bp.route('/reports/organization/thankyou')
@login_required
def organization_thankyou():
    # Get filter parameters
    year = request.args.get('year', datetime.now().year)
    
    # Query organization participation through EventParticipation
    org_stats = db.session.query(
        Organization,
        db.func.count(db.distinct(Event.id)).label('unique_sessions'),
        db.func.sum(EventParticipation.delivery_hours).label('total_hours'),
        db.func.count(db.distinct(Volunteer.id)).label('unique_volunteers')
    ).join(
        VolunteerOrganization, Organization.id == VolunteerOrganization.organization_id
    ).join(
        Volunteer, VolunteerOrganization.volunteer_id == Volunteer.id
    ).join(
        EventParticipation, Volunteer.id == EventParticipation.volunteer_id
    ).join(
        Event, EventParticipation.event_id == Event.id
    ).filter(
        extract('year', Event.start_date) == year,
        EventParticipation.status == 'Attended'
    ).group_by(
        Organization.id
    ).order_by(
        db.desc('total_hours')
    ).all()

    # Format the data for the template
    org_data = [{
        'id': org.id,
        'name': org.name,
        'unique_sessions': sessions,
        'total_hours': round(hours or 0, 2),
        'unique_volunteers': volunteers
    } for org, sessions, hours, volunteers in org_stats]

    return render_template(
        'reports/organization_thankyou.html',
        organizations=org_data,
        year=year,
        now=datetime.now()
    )

@report_bp.route('/reports/organization/thankyou/detail/<int:org_id>')
@login_required
def organization_thankyou_detail(org_id):
    # Get filter parameters
    year = request.args.get('year', datetime.now().year)
    
    # Get organization details
    organization = Organization.query.get_or_404(org_id)
    
    # Query volunteers and their participation for this organization
    volunteer_stats = db.session.query(
        Volunteer,
        db.func.count(db.distinct(Event.id)).label('event_count'),
        db.func.sum(EventParticipation.delivery_hours).label('total_hours')
    ).join(
        VolunteerOrganization, Volunteer.id == VolunteerOrganization.volunteer_id
    ).join(
        EventParticipation, Volunteer.id == EventParticipation.volunteer_id
    ).join(
        Event, EventParticipation.event_id == Event.id
    ).filter(
        VolunteerOrganization.organization_id == org_id,
        extract('year', Event.start_date) == year,
        EventParticipation.status == 'Attended'
    ).group_by(
        Volunteer.id
    ).order_by(
        db.desc('total_hours')
    ).all()

    # Format volunteer data
    volunteers_data = [{
        'id': v.id,
        'name': f"{v.first_name} {v.last_name}",
        'events': events,
        'hours': round(hours or 0, 2)
    } for v, events, hours in volunteer_stats]

    # Get event details
    events = db.session.query(
        Event,
        db.func.count(db.distinct(EventParticipation.volunteer_id)).label('volunteer_count'),
        db.func.sum(EventParticipation.delivery_hours).label('total_hours')
    ).join(
        EventParticipation, Event.id == EventParticipation.event_id
    ).join(
        Volunteer, EventParticipation.volunteer_id == Volunteer.id
    ).join(
        VolunteerOrganization, Volunteer.id == VolunteerOrganization.volunteer_id
    ).filter(
        VolunteerOrganization.organization_id == org_id,
        extract('year', Event.start_date) == year,
        EventParticipation.status == 'Attended'
    ).group_by(
        Event.id
    ).order_by(
        Event.start_date
    ).all()

    # Format event data
    events_data = [{
        'date': event.start_date.strftime('%B %d, %Y'),
        'title': event.title,
        'type': event.type.value if event.type else 'Unknown',
        'volunteers': vol_count,
        'hours': round(hours or 0, 2)
    } for event, vol_count, hours in events]

    return render_template(
        'reports/organization_thankyou_detail.html',
        organization=organization,
        volunteers=volunteers_data,
        events=events_data,
        year=year,
        now=datetime.now()
    )

def get_current_school_year():
    """
    Returns the current school year in 'YYZZ' format (e.g., '2324' for 2023-24).
    If before June 1st, returns current academic year.
    If after June 1st, returns next academic year.
    """
    today = datetime.now()
    if today.month < 6:  # Before June
        return f"{str(today.year-1)[-2:]}{str(today.year)[-2:]}"
    return f"{str(today.year)[-2:]}{str(today.year+1)[-2:]}"

def get_school_year_date_range(school_year):
    """
    Returns start and end dates for a school year.
    school_year format: '2324' for 2023-2024 school year
    """
    year = int(school_year[:2]) + 2000
    start_date = datetime(year, 6, 1)  # June 1st of start year
    end_date = datetime(year + 1, 5, 31)  # May 31st of end year
    return start_date, end_date

@report_bp.route('/reports/district/year-end')
@login_required
def district_year_end():
    # Get school year from query params or default to current
    school_year = request.args.get('school_year', get_current_school_year())
    print(f"Requested school year: {school_year}")  # Debug log
    
    # Get cached reports for the school year
    cached_reports = DistrictYearEndReport.query.filter_by(school_year=school_year).all()
    print(f"Found {len(cached_reports)} cached reports")  # Debug log
    
    district_stats = {report.district.name: report.report_data for report in cached_reports}
    
    if not district_stats:
        print("No cached data found, generating new stats")  # Debug log
        district_stats = generate_district_stats(school_year)
        cache_district_stats(school_year, district_stats)
    
    # Generate list of school years (from 2020-21 to current+1)
    current_year = int(get_current_school_year()[:2])
    school_years = [f"{y}{y+1}" for y in range(20, current_year + 2)]
    school_years.reverse()  # Most recent first
    
    # Convert UTC time to Central time for display
    last_updated = None
    if cached_reports:
        utc_time = min(report.last_updated for report in cached_reports)
        # Convert to Central time (UTC-6 or UTC-5 depending on daylight savings)
        central = pytz.timezone('America/Chicago')
        last_updated = utc_time.replace(tzinfo=pytz.UTC).astimezone(central)
    
    return render_template(
        'reports/district_year_end.html',
        districts=district_stats,
        school_year=school_year,
        school_years=school_years,
        now=datetime.now(),
        last_updated=last_updated
    )

@report_bp.route('/reports/district/year-end/refresh', methods=['POST'])
@login_required
def refresh_district_year_end():
    """Refresh the cached district year-end report data"""
    school_year = request.args.get('school_year', get_current_school_year())
    print(f"Starting refresh for school year: {school_year}")  # Debug log
    
    try:
        # Delete existing cached reports for this school year
        deleted_count = DistrictYearEndReport.query.filter_by(school_year=school_year).delete()
        print(f"Deleted {deleted_count} existing cache entries")  # Debug log
        db.session.commit()
        
        # Generate new stats
        district_stats = generate_district_stats(school_year)
        print(f"Generated new stats for {len(district_stats)} districts")  # Debug log
        
        # Debug log event types for each district
        for district_name, stats in district_stats.items():
            print(f"District {district_name} event types: {stats['event_types']}")
            event_count = sum(stats['event_types'].values())
            print(f"Total events for {district_name}: {event_count}")
        
        # Cache the stats
        cache_district_stats(school_year, district_stats)
        print("Successfully cached new stats")  # Debug log
        
        return jsonify({
            'success': True, 
            'message': f'Successfully refreshed data for {school_year[:2]}-{school_year[2:]} school year'
        })
    except Exception as e:
        db.session.rollback()
        print(f"Error during refresh: {str(e)}")  # Debug log
        return jsonify({'success': False, 'error': str(e)}), 500

def generate_district_stats(school_year):
    """Generate district statistics for a school year"""
    district_stats = {}
    start_date, end_date = get_school_year_date_range(school_year)
    print(f"Date range: {start_date} to {end_date}")  # Debug log
    
    # Only exclude connector sessions
    excluded_event_types = ['connector_session']
    print(f"Excluded event types: {excluded_event_types}")  # Debug log
    
    # Get all districts from the database
    all_districts = District.query.order_by(District.name).all()
    
    # Process each district in our mapping
    for salesforce_id, mapping in DISTRICT_MAPPING.items():
        if not mapping['show']:
            continue
            
        print(f"\nProcessing district: {mapping['name']}")  # Debug log
        
        # Find the primary district record
        primary_district = next((d for d in all_districts if d.salesforce_id == salesforce_id), None)
        if not primary_district:
            print(f"Warning: Primary district {mapping['name']} not found in database")
            continue
            
        # Get all schools for this district
        schools = School.query.filter_by(district_id=primary_district.id).all()
        print(f"Found {len(schools)} schools for district")  # Debug log
        
        # Build the query conditions for this district
        query_conditions = [
            Event.districts.contains(primary_district),
            Event.school.in_([school.id for school in schools])
        ]

        # Add title and district_partner conditions
        title_conditions = [Event.title.ilike(f"%{school.name}%") for school in schools]
        district_partner_conditions = [
            Event.district_partner.ilike(f"%{school.name}%") for school in schools
        ]
        district_partner_conditions.append(Event.district_partner.ilike(f"%{primary_district.name}%"))

        # Add alias conditions
        if 'aliases' in mapping:
            for alias in mapping['aliases']:
                district_partner_conditions.append(Event.district_partner.ilike(f"%{alias}%"))

        # Combine all conditions
        query_conditions.extend([
            db.or_(*title_conditions) if title_conditions else True,
            db.or_(*district_partner_conditions) if district_partner_conditions else True
        ])

        # Query events using the conditions
        events_query = Event.query.filter(
            db.or_(*query_conditions),
            Event.start_date >= start_date,
            Event.start_date <= end_date,
            Event.status == EventStatus.COMPLETED,
            ~Event.type.in_([EventType(t) for t in excluded_event_types])
        )
        
        events = events_query.distinct().all()
        print(f"Found {len(events)} total events")  # Debug log
        
        # Debug log event types
        event_types_found = {}
        for event in events:
            event_type = event.type.value if event.type else 'unknown'
            event_types_found[event_type] = event_types_found.get(event_type, 0) + 1
        print(f"Event types found: {event_types_found}")  # Debug log
        
        # Initialize stats dictionary
        stats = {
            'name': mapping['name'],
            'district_code': mapping['district_code'],
            'total_events': len(events),
            'total_students': 0,
            'total_volunteers': 0,
            'total_volunteer_hours': 0,
            'event_types': event_types_found,
            'schools_reached': set(),
            'monthly_breakdown': {},
            'career_clusters': set()
        }
        
        # Calculate statistics
        for event in events:
            # Get student count based on event type
            student_count = 0
            if event.type == EventType.VIRTUAL_SESSION:
                student_count = event.attended_count or 0
            elif hasattr(event, 'attendance') and event.attendance:
                student_count = event.attendance.total_attendance
            else:
                student_count = event.participant_count or 0
            
            stats['total_students'] += student_count
            
            # Track volunteer participation with improved counting for virtual events
            if event.type == EventType.VIRTUAL_SESSION:
                # For virtual sessions, count all volunteers with Completed or Successfully Completed status
                volunteer_participations = EventParticipation.query.filter(
                    EventParticipation.event_id == event.id,
                    EventParticipation.status.in_(['Attended', 'Completed', 'Successfully Completed'])
                ).all()
                
                # If we found participations, use them
                if volunteer_participations:
                    volunteer_count = len(volunteer_participations)
                    volunteer_hours = sum(p.delivery_hours or 0 for p in volunteer_participations)
                else:
                    # If no participations but the event is completed, assume 1 volunteer and 1 hour
                    volunteer_count = 1
                    volunteer_hours = 1.0
            else:
                # For non-virtual events, use the standard attendance counting
                volunteer_participations = EventParticipation.query.filter(
                    EventParticipation.event_id == event.id,
                    EventParticipation.status.in_(['Attended', 'Completed', 'Successfully Completed'])
                ).all()
                
                volunteer_count = len(volunteer_participations)
                volunteer_hours = sum(p.delivery_hours or 0 for p in volunteer_participations)
            
            # Add debug info for virtual events
            if event.type == EventType.VIRTUAL_SESSION:
                print(f"Virtual event: '{event.title}', Volunteers: {volunteer_count}, Hours: {volunteer_hours}")
            
            stats['total_volunteers'] += volunteer_count
            stats['total_volunteer_hours'] += volunteer_hours
            
            # Track schools and career clusters
            if event.school:
                stats['schools_reached'].add(event.school)
            if event.series:
                stats['career_clusters'].add(event.series)
            
            # Monthly breakdown
            month = event.start_date.strftime('%B %Y')
            if month not in stats['monthly_breakdown']:
                stats['monthly_breakdown'][month] = {
                    'events': 0,
                    'students': 0,
                    'volunteers': 0,
                    'volunteer_hours': 0
                }
            stats['monthly_breakdown'][month]['events'] += 1
            stats['monthly_breakdown'][month]['students'] += student_count
            stats['monthly_breakdown'][month]['volunteers'] += volunteer_count
            stats['monthly_breakdown'][month]['volunteer_hours'] += volunteer_hours
        
        # Convert sets to counts and round hours
        stats['schools_reached'] = len(stats['schools_reached'])
        stats['career_clusters'] = len(stats['career_clusters'])
        stats['total_volunteer_hours'] = round(stats['total_volunteer_hours'], 1)
        
        # Round monthly volunteer hours
        for month in stats['monthly_breakdown']:
            stats['monthly_breakdown'][month]['volunteer_hours'] = round(
                stats['monthly_breakdown'][month]['volunteer_hours'], 1
            )
        
        district_stats[mapping['name']] = stats
    
    return district_stats

def cache_district_stats(school_year, district_stats):
    """Save district statistics to the cache table"""
    for district_name, stats in district_stats.items():
        district = District.query.filter_by(name=district_name).first()
        if district:
            # Update or create cache entry
            report = DistrictYearEndReport.query.filter_by(
                district_id=district.id,
                school_year=school_year  # Changed from year to school_year
            ).first() or DistrictYearEndReport(
                district_id=district.id,
                school_year=school_year  # Changed from year to school_year
            )
            report.report_data = stats
            report.last_updated = datetime.utcnow()
            db.session.add(report)
    
    db.session.commit()

def update_event_districts(event, district_names):
    """Helper function to update event district relationships"""
    # Clear existing relationships
    event.districts = []
    
    for name in district_names:
        # Try exact match first
        district = District.query.filter(District.name.ilike(name)).first()
        if district and district not in event.districts:
            event.districts.append(district)
            
        # Update text field for backward compatibility
        if event.district_partner:
            current_districts = set(d.strip() for d in event.district_partner.split(','))
            current_districts.add(name)
            event.district_partner = ', '.join(current_districts)
        else:
            event.district_partner = name

@report_bp.route('/reports/recruitment')
@login_required
def recruitment_tools():
    recruitment_tools = [
        {
            'title': 'Quick Recruitment Tool',
            'description': 'View upcoming unfilled events and search volunteers by skills and availability.',
            'icon': 'fa-solid fa-bolt',
            'url': '/reports/recruitment/quick',
            'category': 'Recruitment'
        },
        {
            'title': 'Advanced Volunteer Search',
            'description': 'Search volunteers using multiple filters including skills, past events, and volunteer history.',
            'icon': 'fa-solid fa-search',
            'url': '/reports/recruitment/search',
            'category': 'Recruitment'
        }
    ]
    
    return render_template('reports/recruitment_tools.html', tools=recruitment_tools)

# Rename the existing recruitment route to quick_recruitment
@report_bp.route('/reports/recruitment/quick')
@login_required
def quick_recruitment():
    # Get upcoming events that need volunteers from the UpcomingEvent model
    upcoming_events = UpcomingEvent.query.filter(
        UpcomingEvent.start_date >= datetime.now(),
        UpcomingEvent.available_slots > UpcomingEvent.filled_volunteer_jobs
    ).order_by(UpcomingEvent.start_date).all()

    # Get the search query from the request
    search_query = request.args.get('search', '').strip().lower()
    page = request.args.get('page', 1, type=int)
    per_page = 20  # Number of volunteers per page

    # Get the event type filter
    event_type_filter = request.args.get('event_type', '').strip()
    
    # Fix: Handle exclude_dia value from both checkbox and hidden input
    exclude_dia = request.args.get('exclude_dia')
    exclude_dia = exclude_dia == '1' or exclude_dia == 'true' or exclude_dia == 'True'

    # Initialize volunteers_data and pagination as None
    volunteers_data = []
    pagination = None

    # Only query volunteers if there is a search query
    if search_query:
        # Split search query into words
        search_terms = search_query.split()
        
        volunteers_query = Volunteer.query.outerjoin(
            Volunteer.volunteer_organizations
        ).outerjoin(
            VolunteerOrganization.organization
        ).outerjoin(
            EventParticipation, EventParticipation.volunteer_id == Volunteer.id
        ).filter(
            db.or_(
                db.and_(*[
                    db.or_(
                        Volunteer.first_name.ilike(f'%{term}%'),
                        Volunteer.last_name.ilike(f'%{term}%')
                    ) for term in search_terms
                ]),
                Volunteer.title.ilike(f'%{search_query}%'),
                Organization.name.ilike(f'%{search_query}%'),
                Volunteer.skills.any(Skill.name.ilike(f'%{search_query}%'))
            )
        )

        # Add pagination to the volunteers query
        paginated_volunteers = volunteers_query.add_columns(
            db.func.count(EventParticipation.id).label('participation_count'),
            db.func.max(EventParticipation.event_id).label('last_event')
        ).group_by(Volunteer.id).paginate(page=page, per_page=per_page, error_out=False)

        # Store pagination object for template
        pagination = paginated_volunteers

        # Populate volunteers_data based on the paginated results
        for volunteer, participation_count, last_event in paginated_volunteers.items:
            volunteers_data.append({
                'id': volunteer.id,
                'name': f"{volunteer.first_name} {volunteer.last_name}",
                'email': volunteer.primary_email,
                'title': volunteer.title,
                'organization': {
                    'name': volunteer.volunteer_organizations[0].organization.name if volunteer.volunteer_organizations else None,
                    'id': volunteer.volunteer_organizations[0].organization.id if volunteer.volunteer_organizations else None
                },
                'participation_count': participation_count,
                'skills': [skill.name for skill in volunteer.skills],
                'industry': volunteer.industry,
                'last_mailchimp_date': volunteer.last_mailchimp_activity_date,
                'last_volunteer_date': volunteer.last_volunteer_date,
                'last_email_date': volunteer.last_email_date,
                'salesforce_contact_url': volunteer.salesforce_contact_url
            })

    # Prepare events_data based on UpcomingEvent model
    events_data = []
    for event in upcoming_events:
        # Extract the URL from the HTML anchor tag if it exists
        registration_link = None
        if event.registration_link and 'href=' in event.registration_link:
            start = event.registration_link.find('href="') + 6
            end = event.registration_link.find('"', start)
            if start > 5 and end > start:  # ensure we found both quotes
                registration_link = event.registration_link[start:end]
        
        events_data.append({
            'title': event.name,
            'description': event.event_type,
            'start_date': event.start_date,
            'type': event.event_type,
            'location': event.registration_link,
            'total_slots': event.available_slots,
            'filled_slots': event.filled_volunteer_jobs,
            'remaining_slots': event.available_slots - event.filled_volunteer_jobs,
            'skills_needed': [],
            'status': 'Upcoming',
            'registration_link': registration_link
        })

    return render_template(
        'reports/recruitment_report.html',
        events=events_data,
        volunteers=volunteers_data,
        search_query=search_query,
        event_types=UpcomingEvent.query.with_entities(UpcomingEvent.event_type).distinct().all(),
        exclude_dia=exclude_dia,
        event_type_filter=event_type_filter,
        pagination=pagination,
        page=page
    )

@report_bp.route('/reports/contact')
@login_required
def contact_report():
    # Get search parameters
    search = request.args.get('search', '').strip()
    sort = request.args.get('sort', 'start_date')
    order = request.args.get('order', 'asc')

    # Base query for future events that are confirmed
    query = Event.query.filter(
        Event.start_date >= datetime.utcnow(),  # Only future events
        Event.status == EventStatus.CONFIRMED
    )
    
    # Apply search if provided
    if search:
        query = query.filter(Event.title.ilike(f'%{search}%'))
    
    # Apply sorting
    if sort == 'title':
        query = query.order_by(Event.title.asc() if order == 'asc' else Event.title.desc())
    elif sort == 'location':
        query = query.order_by(Event.location.asc() if order == 'asc' else Event.location.desc())
    else:  # default to date
        query = query.order_by(Event.start_date.asc() if order == 'asc' else Event.start_date.desc())
    
    upcoming_events = query.all()
    
    # Get participant counts for each event
    event_stats = {}
    for event in upcoming_events:
        participations = event.volunteer_participations
        event_stats[event.id] = {
            'volunteer_count': len(participations),
        }
    
    return render_template('reports/contact_report.html', 
                         events=upcoming_events,
                         event_stats=event_stats,
                         search=search,
                         sort=sort,
                         order=order)

@report_bp.route('/reports/contact/<int:event_id>')
@login_required
def contact_report_detail(event_id):
    event = Event.query.get_or_404(event_id)
    
    # Get sort parameters
    sort = request.args.get('sort', 'name')
    order = request.args.get('order', 'asc')
    
    # Get all participants for this event
    participations = event.volunteer_participations
    
    # Group participants by status
    participants_by_status = {
        'Scheduled': [],
        'Unscheduled': []
    }
    
    for participation in participations:
        if participation.status == 'Scheduled':
            participants_by_status['Scheduled'].append(participation)
        else:
            participants_by_status['Unscheduled'].append(participation)
    
    # Sort participants in each group
    for status in participants_by_status:
        participants = participants_by_status[status]
        if sort == 'name':
            participants.sort(
                key=lambda p: f"{p.volunteer.first_name} {p.volunteer.last_name}",
                reverse=(order == 'desc')
            )
        elif sort == 'title':
            participants.sort(
                key=lambda p: p.title or p.volunteer.title or '',
                reverse=(order == 'desc')
            )
        elif sort == 'email':
            participants.sort(
                key=lambda p: p.volunteer.primary_email or '',
                reverse=(order == 'desc')
            )
        elif sort == 'organization':
            participants.sort(
                key=lambda p: p.volunteer.organizations[0].name if p.volunteer.organizations else '',
                reverse=(order == 'desc')
            )
    
    return render_template('reports/contact_report_detail.html',
                         event=event,
                         participants_by_status=participants_by_status,
                         sort=sort,
                         order=order)

@report_bp.route('/reports/district/year-end/<district_name>')
@login_required
def district_year_end_detail(district_name):
    """Show detailed year-end report for a specific district"""
    school_year = request.args.get('school_year', get_current_school_year())
    
    # Get district
    district = District.query.filter_by(name=district_name).first_or_404()
    
    # Get the district's mapping info
    district_mapping = next((mapping for salesforce_id, mapping in DISTRICT_MAPPING.items() 
                            if mapping['name'] == district_name), None)
    
    # Try to get cached data first
    cached_report = DistrictYearEndReport.query.filter_by(
        district_id=district.id,
        school_year=school_year
    ).first()
    
    if cached_report:
        stats = cached_report.report_data
    else:
        # Generate stats for just this district
        stats = generate_district_stats(school_year)[district_name]
    
    # Get all events for this district within the school year
    start_date, end_date = get_school_year_date_range(school_year)
    
    # Define event types to exclude - remove virtual_session from exclusion
    excluded_event_types = ['connector_session']

    # Base conditions
    query_conditions = [
        Event.districts.contains(district),
        Event.school.in_([school.id for school in district.schools]),
        *[Event.title.ilike(f"%{school.name}%") for school in district.schools],
        *[Event.district_partner.ilike(f"%{school.name}%") for school in district.schools],
        Event.district_partner.ilike(f"%{district.name}%"),
        Event.district_partner.ilike(f"%{district.name.replace(' School District', '')}%")
    ]

    # Add alias conditions if they exist
    if district_mapping and 'aliases' in district_mapping:
        for alias in district_mapping['aliases']:
            query_conditions.append(Event.district_partner.ilike(f"%{alias}%"))

    # Update the query to include EventAttendance and status filter
    events = (Event.query
        .outerjoin(School, Event.school == School.id)
        .outerjoin(EventAttendance, Event.id == EventAttendance.event_id)
        .filter(
            Event.start_date.between(start_date, end_date),
            Event.status == EventStatus.COMPLETED,
            ~Event.type.in_([EventType(t) for t in excluded_event_types]),
            db.or_(*query_conditions)
        )
        .order_by(Event.start_date)
        .all())

    # Organize events by month
    events_by_month = {}
    for event in events:
        month = event.start_date.strftime('%B %Y')
        if month not in events_by_month:
            events_by_month[month] = {
                'events': [],
                'total_students': 0,
                'total_volunteers': 0,
                'total_volunteer_hours': 0
            }
        
        # Get attendance count from EventAttendance if available, or attended_count for virtual sessions
        student_count = 0
        if event.type == EventType.VIRTUAL_SESSION:
            student_count = event.attended_count or 0
        elif hasattr(event, 'attendance') and event.attendance:
            student_count = event.attendance.total_attendance
        else:
            student_count = event.participant_count or 0
            
        # Get volunteer stats for this event with special handling for virtual events
        if event.type == EventType.VIRTUAL_SESSION:
            # For virtual sessions, count all volunteers with Completed or Successfully Completed status
            volunteer_participations = [p for p in event.volunteer_participations 
                                      if p.status in ['Attended', 'Completed', 'Successfully Completed']]
            
            if volunteer_participations:
                volunteer_count = len(volunteer_participations)
                volunteer_hours = sum(p.delivery_hours or 0 for p in volunteer_participations)
            else:
                # If no participations but the event is completed, assume 1 volunteer and 1 hour
                volunteer_count = 1
                volunteer_hours = 1.0
        else:
            # For non-virtual events, use the standard attendance counting
            volunteer_participations = EventParticipation.query.filter(
                EventParticipation.event_id == event.id,
                EventParticipation.status.in_(['Attended', 'Completed', 'Successfully Completed'])
            ).all()
            
            volunteer_count = len(volunteer_participations)
            volunteer_hours = sum(p.delivery_hours or 0 for p in volunteer_participations)
        
        # Get school name safely
        school_name = None
        if hasattr(event, 'school') and event.school is not None:
            if isinstance(event.school, School):
                school_name = event.school.name
            elif isinstance(event.school, str):
                school = School.query.get(event.school)
                school_name = school.name if school else None
        
        events_by_month[month]['events'].append({
            'id': event.id,
            'title': event.title,
            'date': event.start_date.strftime('%m/%d/%Y'),
            'time': event.start_date.strftime('%I:%M %p'),
            'type': event.type.value if event.type else 'Unknown',
            'students': student_count,
            'volunteers': volunteer_count,
            'volunteer_hours': round(volunteer_hours, 1),
            'school': school_name,
            'location': event.location or 'Virtual'
        })
        
        events_by_month[month]['total_students'] += student_count
        events_by_month[month]['total_volunteers'] += volunteer_count
        events_by_month[month]['total_volunteer_hours'] += volunteer_hours

    # Organize events by school
    schools_by_level = {
        'High': [],
        'Middle': [],
        'Elementary': [],
        None: []  # For schools without a level
    }
    
    # Get all schools in the district
    district_schools = School.query.filter_by(district_id=district.id).all()
    
    # Create a mapping of school IDs to their data
    school_data = {}
    for school in district_schools:
        school_data[school.id] = {
            'name': school.name,
            'level': school.level,
            'events': [],
            'total_students': 0,
            'total_volunteers': 0,
            'total_volunteer_hours': 0
        }
    
    # Process events for each school
    for event in events:
        school_id = event.school if isinstance(event.school, str) else None
        if school_id in school_data:
            school = school_data[school_id]
            
            # Get attendance and volunteer data
            student_count = event.participant_count or 0
            volunteer_count = len([p for p in event.volunteer_participations if p.status == 'Attended'])
            volunteer_hours = sum([p.delivery_hours or 0 for p in event.volunteer_participations if p.status == 'Attended'])
            
            # Add event data - Add type to the event data
            school['events'].append({
                'id': event.id,
                'title': event.title,
                'date': event.start_date.strftime('%m/%d/%Y'),
                'type': event.type.value if event.type else 'Unknown',
                'students': student_count,
                'volunteers': volunteer_count,
                'volunteer_hours': volunteer_hours
            })
            
            # Update totals
            school['total_students'] += student_count
            school['total_volunteers'] += volunteer_count
            school['total_volunteer_hours'] += volunteer_hours
    
    # Organize schools by level
    for school_id, data in school_data.items():
        school = School.query.get(school_id)
        level = school.level if school else None
        if level not in schools_by_level:
            level = None
        schools_by_level[level].append(data)
    
    # Sort schools within each level by name
    for level in schools_by_level:
        schools_by_level[level].sort(key=lambda x: x['name'])

    return render_template(
        'reports/district_year_end_detail.html',
        district=district,
        school_year=school_year,
        stats=stats,
        events_by_month=events_by_month,
        schools_by_level=schools_by_level,
        total_events=len(events)
    )

@report_bp.route('/reports/recruitment/search')
@login_required
def recruitment_search():
    # Get search query, pagination, and sorting parameters
    search_query = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_by = request.args.get('sort', 'name')
    order = request.args.get('order', 'asc')
    search_mode = request.args.get('search_mode', 'wide')  # Default to wide search
    
    # Base query joining necessary tables
    query = Volunteer.query.outerjoin(
        VolunteerOrganization
    ).outerjoin(
        Organization
    ).outerjoin(
        VolunteerSkill
    ).outerjoin(
        Skill
    ).outerjoin(
        EventParticipation
    ).outerjoin(
        Event
    )
    
    # Apply search if provided
    if search_query:
        # Split search query into terms and remove empty strings
        search_terms = [term.strip() for term in search_query.split() if term.strip()]
        
        if search_mode == 'wide':
            # OR mode: Match any term across all fields
            search_conditions = []
            for term in search_terms:
                search_conditions.append(
                    db.or_(
                        Volunteer.first_name.ilike(f'%{term}%'),
                        Volunteer.last_name.ilike(f'%{term}%'),
                        Organization.name.ilike(f'%{term}%'),
                        Skill.name.ilike(f'%{term}%'),
                        Event.title.ilike(f'%{term}%')
                    )
                )
            query = query.filter(db.or_(*search_conditions))
        else:
            # Narrow mode: Must match all terms
            for term in search_terms:
                query = query.filter(
                    db.or_(
                        Volunteer.first_name.ilike(f'%{term}%'),
                        Volunteer.last_name.ilike(f'%{term}%'),
                        Organization.name.ilike(f'%{term}%'),
                        Skill.name.ilike(f'%{term}%'),
                        Event.title.ilike(f'%{term}%')
                    )
                )
    
    # Apply sorting
    if sort_by == 'name':
        if order == 'asc':
            query = query.order_by(Volunteer.first_name, Volunteer.last_name)
        else:
            query = query.order_by(db.desc(Volunteer.first_name), db.desc(Volunteer.last_name))
    elif sort_by == 'organization':
        if order == 'asc':
            query = query.order_by(Organization.name)
        else:
            query = query.order_by(db.desc(Organization.name))
    elif sort_by == 'last_email':
        if order == 'asc':
            query = query.order_by(Volunteer.last_non_internal_email_date)
        else:
            query = query.order_by(db.desc(Volunteer.last_non_internal_email_date))
    elif sort_by == 'last_volunteer':
        if order == 'asc':
            query = query.order_by(Volunteer.last_volunteer_date)
        else:
            query = query.order_by(db.desc(Volunteer.last_volunteer_date))
    elif sort_by == 'times_volunteered':
        subquery = db.session.query(
            EventParticipation.volunteer_id,
            db.func.count(EventParticipation.id).label('volunteer_count')
        ).group_by(EventParticipation.volunteer_id).subquery()
        
        query = query.outerjoin(
            subquery, Volunteer.id == subquery.c.volunteer_id
        ).order_by(
            db.desc(subquery.c.volunteer_count) if order == 'desc' else subquery.c.volunteer_count
        )
    
    # Add pagination
    pagination = query.distinct().paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )
    
    return render_template(
        'reports/recruitment_search.html',
        volunteers=pagination.items,
        search_query=search_query,
        pagination=pagination,
        sort_by=sort_by,
        order=order,
        search_mode=search_mode  # Pass the search mode to the template
    )

@report_bp.route('/reports/pathways')
@login_required
def pathways_report():
    # Get filter parameters
    search = request.args.get('search', '').strip()
    selected_year = int(request.args.get('year', datetime.now().year))
    
    # Base query
    query = Pathway.query
    
    # Apply search filter if provided
    if search:
        query = query.filter(Pathway.name.ilike(f'%{search}%'))
    
    # Get all pathways
    pathways = query.all()
    
    # Get current year for the year filter dropdown
    current_year = datetime.now().year
    
    return render_template('reports/pathways.html',
                         pathways=pathways,
                         search=search,
                         selected_year=selected_year,
                         current_year=current_year)

@report_bp.route('/reports/pathways/<int:pathway_id>')
@login_required
def pathway_detail(pathway_id):
    # Get the pathway
    pathway = Pathway.query.get_or_404(pathway_id)
    
    # Get active events (not completed or cancelled)
    active_events = [event for event in pathway.events 
                    if event.status not in [EventStatus.COMPLETED, EventStatus.CANCELLED]]
    
    # Calculate total attendance
    total_attendance = sum(event.attended_count or 0 for event in pathway.events)
    
    return render_template('reports/pathway_detail.html',
                         pathway=pathway,
                         active_events=active_events,
                         total_attendance=total_attendance)

@report_bp.route('/reports/district/year-end/<district_name>/excel')
@login_required
def district_year_end_excel(district_name):
    """Generate Excel file for district year-end report"""
    school_year = request.args.get('school_year', get_current_school_year())
    
    # Get district
    district = District.query.filter_by(name=district_name).first_or_404()
    
    # Get date range for school year
    start_date, end_date = get_school_year_date_range(school_year)
    
    # Define event types to exclude
    excluded_event_types = ['connector_session']

    # Query events for this district in the given school year
    events = (Event.query
        .outerjoin(School, Event.school == School.id)
        .outerjoin(EventAttendance, Event.id == EventAttendance.event_id)
        .filter(
            Event.start_date.between(start_date, end_date),
            Event.status == EventStatus.COMPLETED,
            ~Event.type.in_([EventType(t) for t in excluded_event_types]),
            db.or_(
                Event.districts.contains(district),
                Event.school.in_([school.id for school in district.schools]),
                *[Event.title.ilike(f"%{school.name}%") for school in district.schools],
                *[Event.district_partner.ilike(f"%{school.name}%") for school in district.schools],
                Event.district_partner.ilike(f"%{district.name}%"),
                Event.district_partner.ilike(f"%{district.name.replace(' School District', '')}%")
            )
        )
        .order_by(Event.start_date)
        .all())
    
    # Prepare data for Excel
    data = []
    for event in events:
        # Get student count
        student_count = 0
        if event.type == EventType.VIRTUAL_SESSION:
            student_count = event.attended_count or 0
        elif hasattr(event, 'attendance') and event.attendance:
            student_count = event.attendance.total_attendance
        else:
            student_count = event.participant_count or 0
            
        # Get volunteer stats
        if event.type == EventType.VIRTUAL_SESSION:
            volunteer_participations = [p for p in event.volunteer_participations 
                                      if p.status in ['Attended', 'Completed', 'Successfully Completed']]
            
            if volunteer_participations:
                volunteer_count = len(volunteer_participations)
                volunteer_hours = sum(p.delivery_hours or 0 for p in volunteer_participations)
            else:
                # If no participations but event is completed, assume 1 volunteer and 1 hour
                volunteer_count = 1
                volunteer_hours = 1.0
        else:
            volunteer_count = len([p for p in event.volunteer_participations if p.status == 'Attended'])
            volunteer_hours = sum([p.delivery_hours or 0 for p in event.volunteer_participations if p.status == 'Attended'])
        
        # Get school info
        school_name = None
        school_level = None
        if hasattr(event, 'school') and event.school is not None:
            if isinstance(event.school, School):
                school_name = event.school.name
                school_level = event.school.level
            elif isinstance(event.school, str):
                school = School.query.get(event.school)
                if school:
                    school_name = school.name
                    school_level = school.level
        
        # Add event data
        data.append({
            'Date': event.start_date.strftime('%m/%d/%Y'),
            'Time': event.start_date.strftime('%I:%M %p'),
            'Event': event.title,
            'Type': event.type.value if event.type else 'Unknown',
            'Location': event.location or 'Virtual',
            'Students': student_count,
            'Volunteers': volunteer_count,
            'Hours': round(volunteer_hours, 1),
            'School': school_name or 'N/A',
            'Level': school_level or 'N/A'
        })
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Write data to Excel
    df.to_excel(writer, sheet_name='Events', index=False)
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Events']
    
    # Add some formatting
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'bg_color': '#D6E9F8',
        'border': 1
    })
    
    # Write the column headers with the defined format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    # Set column widths
    worksheet.set_column('A:A', 12)  # Date
    worksheet.set_column('B:B', 10)  # Time
    worksheet.set_column('C:C', 30)  # Event
    worksheet.set_column('D:D', 15)  # Type
    worksheet.set_column('E:E', 20)  # Location
    worksheet.set_column('F:F', 10)  # Students
    worksheet.set_column('G:G', 10)  # Volunteers
    worksheet.set_column('H:H', 10)  # Hours
    worksheet.set_column('I:I', 25)  # School
    worksheet.set_column('J:J', 15)  # Level
    
    # Close the writer and output the Excel file
    writer.close()
    output.seek(0)
    
    # Generate filename
    filename = f"{district_name.replace(' ', '_')}_{school_year}_Year_End_Report.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name=filename,
        as_attachment=True
    )