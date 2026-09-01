import os
import re
import html
from datetime import datetime
from hmpps.services.job_log_handling import (
  log_debug,
  log_error,
  log_warning,
  log_info,
)

log_level = os.environ.get('LOG_LEVEL', 'INFO').upper()


def clean_value(value):
  if value is None:
    return None
  if isinstance(value, str):
    return html.unescape(value).strip()
  return value


def fetchID(sp_product, dict, key):
  if key in sp_product and sp_product[key] is not None:
    parent_key = sp_product[key]
    if parent_key in dict:
      sp_product[key] = dict[parent_key]['documentId']
    else:
      log_error(
        f'Product reference key not found for {key} in Service Catalogue :: '
        f'{sp_product[key]}'
      )
      del sp_product[key]
  return sp_product


# Look for related data in other Sharepoint records
def link_product_data(sp, sp_product):
  log_debug('Linking product with other Sharepoint data')
  product_id = sp_product.get('fields', {}).get('ProductID', None)
  product_data = {}
  # This is a tuple of
  # dictionary_key,lookup_key,sharepoint_list,field_to_return
  fields = [
    ('parent', 'ParentProductLookupId', 'Products and Teams Main List', 'Product'),
    ('team', 'TeamLookupId', 'Teams', 'Team'),
    ('product_set', 'ProductSetLookupId', 'Product Set', 'ProductSet'),
    ('service_area', 'ServiceAreaLookupId', 'Service Areas', 'ServiceArea'),
    (
      'delivery_manager',
      'DeliveryManagerLookupId',
      'Delivery Managers',
      'DeliveryManagerName',
    ),
    (
      'product_manager',
      'ProductManagerLookupId',
      'Product Managers',
      'ProductManagerName',
    ),
    ('lead_developer', 'LeadDeveloperLookupId', 'Lead Developers', 'Title'),
    (
      'technical_architect',
      'TechnicalArchitectLookupId',
      'Technical Architects',
      'TechnicalArchitectName',
    ),
    (
      'principal_architect',
      'OversightPrincipalTechnicalArchiLookupId',
      'Principal Technical Architect',
      'PrincipalTechnicalArchitectName',
    ),
  ]
  for field in fields:
    if field_id := sp_product.get('fields', {}).get(field[1]):
      product_data[field[0]] = (
        sp.dict[field[2]].get(field_id, {}).get('fields', {}).get(field[3], None)
      )
      if not product_data[field[0]]:
        log_warning(
          f'{field[0]} matching {field[1]} not found for product_id: {product_id}'
        )
  return product_data


def extract_sp_products_data(sp):
  sp_products_data = []
  for sp_product in sp.data['Products and Teams Main List'].get('value'):
    log_debug(
      'Extracting SharePoint product data for: '
      f'{sp_product.get("fields", {}).get("ProductID", None)}'
    )
    sp_product_fields = sp_product.get('fields')
    if not sp_product_fields or not isinstance(sp_product_fields, dict):
      log_warning('Invalid product data returned -  ignoring it.')
      continue

    if product_id := sp_product_fields.get('ProductID'):
      product_name = clean_value(sp_product_fields.get('Product')) or ''

      # Check for a valid (non-empty) product name - otherwise skip it
      if not product_name:
        log_warning(
          f'Invalid name format for product_id: {product_id} ({product_name}'
          f' - ignoring it.'
        )
        continue

      # Check for a valid (alphanumeric) Product ID - otherwise skip it
      if not re.match(r'^[A-Z]{3,5}[0-9]{0,5}$', product_id):
        log_warning(
          f'Invalid ProductID format for product_id: {product_id} - ignoring it.'
        )
        continue

      # set subproductBool directly from the comparison
      subproductBool = (
        str(sp_product_fields.get('ProductType', '')).strip().lower() == 'subproduct'
      )

      # fetch links to other Sharepoint lists
      linked_product_data = link_product_data(sp, sp_product)

      sp_product_data = {
        'p_id': product_id,
        'name': product_name,
        'subproduct': subproductBool,
        'description': clean_value(
          sp_product_fields.get('Description_x0028_SourceData_x00', None)
        ),
        'phase': sp_product_fields.get('field_7', None),
        'slack_channel_id': sp_product.get('fields', {}).get('SlackchannelID', None),
        'portfolio': clean_value(sp_product.get('fields', {}).get('Portfolio', None)),
        'business_owner': clean_value(
          sp_product_fields.get('HMPPSBusinessOwner', None)
        ),
        'decommissioned': clean_value(
          str(sp_product_fields.get('DecommissionedProduct', '')).strip().lower()
          == 'yes'
        ),
        'decommissioned_date': format_date(
          sp_product_fields.get('DecommissionedEndDate', None)
        ),
        # "updated_by_id": 34
      }
      # add the fetched data
      sp_product_data.update(linked_product_data)
      sp_products_data.append(sp_product_data)
  return sp_products_data


def format_date(date_str):
  if not date_str:
    return None
  try:
    # Parse the date string and convert to 'DD/MM/YYYY'
    return datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%SZ').strftime('%Y-%m-%d')
  except ValueError:
    log_error(f'Invalid date format: {date_str}')
    return None


def process_sc_products(services):
  def log_and_append(message):
    log_info(message)
    log_messages.append(message)

  sc = services.sc
  sp = services.sp

  # Service Catalogue
  log_info('Processing Products ')

  log_debug(
    'Fetching Products, Teams, Products Sets, Service Areas from Service Catalogue'
  )
  sc_products_data = sc.get_all_records(sc.sharepoint_discovery_products_get)
  sc_teams_data = sc.get_all_records('teams')
  sc_product_sets_data = sc.get_all_records('product-sets')
  sc_service_areas_data = sc.get_all_records('service-areas')
  # Create the dictionaries
  sc_products_dict = {
    product.get('p_id').strip(): product for product in sc_products_data
  }
  sc_product_name_dict = {
    product.get('name').strip(): product for product in sc_products_data
  }
  sc_team_name_dict = {team.get('name').strip(): team for team in sc_teams_data}
  sc_product_set_name_dict = {
    product_set.get('name').strip(): product_set for product_set in sc_product_sets_data
  }
  sc_service_area_name_dict = {
    service_area.get('name').strip(): service_area
    for service_area in sc_service_areas_data
  }

  # Sharepoint data processing
  sp_products_data = extract_sp_products_data(sp)

  # Quick summary before we start
  log_info(f'Found {len(sp_products_data)} products in SharePoint (after processing)')
  log_info(f'Found {len(sc_products_data)} products in Service Catalogue')

  # Compare and update sp_product_data
  log_info('Processing prepared products sharepoint data for service catalogue ')
  change_count = 0
  log_messages = []
  log_info('************** Processing Products *********************')
  for sp_product in sp_products_data:
    p_id = sp_product.get('p_id')
    log_debug(f'Comparing Product p_id {p_id} :: {sp_product}')
    if p_id in sc_products_dict:
      try:
        sc_product = sc_products_dict.get(p_id, {})
        log_debug(
          f'\nComparing SC product {sc_product} \n with SP product {sp_product}'
        )
        mismatch_flag = False
        for key in list(sp_product.keys()):
          sp_value = clean_value(sp_product.get(key))
          sc_value = None
          compare_flag = False
          if key in sp_product and key in sc_product:
            compare_flag = True
          if (
            compare_flag
            and key != 'updated_by_id'
            and key != 'subproduct'
            and key != 'decommissioned'
            and key != 'p_id'
          ):
            if (
              key == 'parent'
              or key == 'team'
              or key == 'product_set'
              or key == 'service_area'
            ):
              if sc_product.get(key):
                try:
                  sc_value = clean_value(sc_product.get(key).get('name'))
                except KeyError:
                  log_error(
                    f'Key {key} not found in Service Catalogue data for p_id {p_id}'
                  )
            else:
              try:
                sc_value = clean_value(sc_product.get(key))
              except KeyError:
                log_error(
                  f'Key {key} not found in Service Catalogue data for p_id {p_id}'
                )

            if sp_value is not None:
              if (sp_value or '').strip() != (sc_value or '').strip():
                log_and_append(
                  f'SC Updating Products p_id {p_id}({key}) :: {sc_value} -> {sp_value}'
                )
                log_info(
                  f'SC Updating Products p_id {p_id}({key}) :: {sc_value} -> {sp_value}'
                )
                mismatch_flag = True
              else:
                del sp_product[key]

          elif compare_flag and key == 'subproduct':
            if sp_product.get(key) != sc_product.get(key):
              log_and_append(
                f'Updating Products p_id {p_id}({key}) :: {sp_value} -> {sc_value}'
              )
              mismatch_flag = True
            else:
              del sp_product[key]
          elif compare_flag and key == 'decommissioned':
            sp_value = sp_product.get(key)
            sc_value = sc_product.get(key) if sc_product.get(key) is not None else False
            if sp_value != sc_value:
              log_and_append(
                f'Updating Products p_id {p_id}({key}) :: {sp_value} -> {sc_value}'
              )
              mismatch_flag = True
            else:
              del sp_product[key]

        if mismatch_flag:
          sp_product = (
            fetchID(sp_product, sc_product_name_dict, 'parent')
            if 'parent' in sp_product
            else sp_product
          )
          sp_product = (
            fetchID(sp_product, sc_team_name_dict, 'team')
            if 'team' in sp_product
            else sp_product
          )
          sp_product = (
            fetchID(sp_product, sc_product_set_name_dict, 'product_set')
            if 'product_set' in sp_product
            else sp_product
          )
          sp_product = (
            fetchID(sp_product, sc_service_area_name_dict, 'service_area')
            if 'service_area' in sp_product
            else sp_product
          )
          sc.update('products', sc_product.get('documentId'), sp_product)
          change_count += 1
      except Exception as e:
        log_error(f'Error processing product p_id {p_id}: {e}')
    else:
      sp_product = fetchID(sp_product, sc_product_name_dict, 'parent')
      sp_product = fetchID(sp_product, sc_team_name_dict, 'team')
      sp_product = fetchID(sp_product, sc_product_set_name_dict, 'product_set')
      sp_product = fetchID(sp_product, sc_service_area_name_dict, 'service_area')
      log_and_append(f'Adding Product :: {sp_product}')
      sc.add('products', sp_product)
      change_count += 1

  log_and_append(f'Products in Service Catalogue processed: {change_count}')
  return log_messages
