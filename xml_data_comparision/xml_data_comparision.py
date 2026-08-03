import os
import xml.etree.ElementTree as ET
import difflib # For string similarity comparison
import csv
import sys      # --- NEW: Added for exit codes and stderr ---
import argparse # --- NEW: Added for command-line arguments ---
import json     # --- NEW: Added for parsing points map ---

try:
    import pandas as pd
    import openpyxl # Used by pandas ExcelWriter engine
except ImportError:
    # --- Send errors to stderr ---
    print("Error: 'pandas' and 'openpyxl' libraries are required. Please install them (`pip install pandas openpyxl`)", file=sys.stderr)
    sys.exit(1) # Exit with error code

from typing import Union, List, Dict, Tuple


# Defining a small tolerance for floating-point comparisons
FLOAT_TOLERANCE = 1e-6

# Threshold for filename similarity (0.5 for 50% match)
FILENAME_SIMILARITY_THRESHOLD = 0.5

# Threshold for feature element similarity (0.85 for 85% match)
# This determines if a student feature / geometry element is considered a "match" for a reference feature.
FEATURE_MATCH_THRESHOLD = 0.85

#Conversipn of all German value representation (commas) into decimals
def normalize_value(value: str) -> Union[float, str]:
    """
    Normalizes a string value for comparison.
    - Replaces comma decimals with dot decimals.
    - Tries to convert to float for numerical comparison.
    - Returns string if not a valid float.
    """
    if isinstance(value, (int, float)):
        return float(value)

    s_value = str(value).strip()
    if ',' in s_value:
        s_value = s_value.replace(',', '.')

    try:
        return float(s_value)
    except ValueError:
        return s_value

#Calculates the similarity between the two XML filenames to be compared
def calculate_filename_similarity(name1: str, name2: str) -> float:
    """
    Calculates the similarity between two filenames as a ratio (0.0 to 1.0).
    Uses difflib.SequenceMatcher to find the longest common subsequence.
    """
    s = difflib.SequenceMatcher(None, name1.lower(), name2.lower())
    return s.ratio()

def calculate_element_similarity(elem1: ET.Element, elem2: ET.Element) -> float:
    """
    Calculates a similarity score between two XML elements.
    This is used to find the best matching feature in an unordered list.
    The score is based on tag, attributes (especially 'Type'), text, and child tags.
    """
    score = 0.0
    max_score = 0.0

    if elem1.tag != elem2.tag:
        return 0.0
    score += 100 # Base score for matching tag
    max_score += 100

    attrs1 = dict(elem1.items())
    attrs2 = dict(elem2.items())

    if 'Type' in attrs1 and 'Type' in attrs2:
        max_score += 50
        if attrs1['Type'] == attrs2['Type']:
            score += 50

    common_attrs = set(attrs1.keys()).intersection(set(attrs2.keys()))
    max_score += len(attrs1) * 5

    for attr_name in common_attrs:
        if attr_name == 'Type' and 'Type' in attrs1 and 'Type' in attrs2:
            continue

        val1 = normalize_value(attrs1[attr_name])
        val2 = normalize_value(attrs2[attr_name])

        if isinstance(val1, float) and isinstance(val2, float):
            if abs(val1 - val2) < FLOAT_TOLERANCE:
                score += 5
        elif val1 == val2:
            score += 5

    text1 = elem1.text.strip() if elem1.text else ""
    text2 = elem2.text.strip() if elem2.text else ""

    normalized_text1 = normalize_value(text1)
    normalized_text2 = normalize_value(text2)

    max_score += 20 # Weight for text content
    if isinstance(normalized_text1, float) and isinstance(normalized_text2, float):
        if abs(normalized_text1 - normalized_text2) < FLOAT_TOLERANCE:
            score += 20
    elif normalized_text1 == normalized_text2:
        score += 20
    else: # Partial match for text if they are strings but not identical
        s = difflib.SequenceMatcher(None, text1.lower(), text2.lower())
        score += 20 * s.ratio() # Add a proportional score based on text similarity

    children1 = list(elem1)
    children2 = list(elem2)

    # Calculate child similarity recursively and add to score
    matched_children_indices = set()
    for child1 in children1:
        best_child_match_score = 0
        best_child_match_index = -1
        for i, child2 in enumerate(children2):
            if i not in matched_children_indices:
                child_sim = calculate_element_similarity(child1, child2)
                if child_sim > best_child_match_score:
                    best_child_match_score = child_sim
                    best_child_match_index = i
        if best_child_match_score > 0: # Consider a match if similarity is greater than 0
            score += best_child_match_score * 10 # Scale child similarity
            if best_child_match_index != -1: # Prevent adding -1 if no match found
                 matched_children_indices.add(best_child_match_index)

    max_score += len(children1) * 10 # Add potential score for all children

    if max_score == 0:
        # If both elements are truly empty (no tags, attrs, text, children considered), they match.
        # Check if tags actually matched initially.
        return 1.0 if elem1.tag == elem2.tag else 0.0

    final_score = min(score / max_score, 1.0) # Ensure score doesn't exceed 1.0
    return final_score


def calculate_part_similarity(comparison_results: List[str]) -> float:
    """
    [Sr01, Sr03]
    Calculates a single part-level similarity score based on the detailed comparison results.
    This implementation counts 'Match' results vs. total results for simplicity.
    Can be expanded for more sophisticated metrics (e.g., weighting different types of matches/mismatches).
    """
    total_results = len(comparison_results)
    if total_results == 0:
         # If no comparisons happened (e.g., root tag mismatch prevented further checks)
         # the similarity should be 0 unless the root tags actually matched implicitly.
         # A more robust check might be needed depending on compare_xml_elements behavior.
         return 0.0 # Changed from 1.0 to reflect potential root mismatch

    match_count = 0
    significant_comparisons = 0 # Count only meaningful comparisons

    for result in comparison_results:
        # Count matches based on your criteria (e.g., specific tags, attributes)
        if "Match" in result:
            # More specific check: ignore simple container matches unless they contain details
            if "Attribute '" in result or "Text :" in result or "(Matched by similarity" in result:
                 match_count += 1
                 significant_comparisons += 1
            # Could add checks for specific feature tags if needed
        # Count significant mismatches/missing/extra as comparisons contributing to the denominator
        elif "Mismatch" in result or "Missing" in result or "Extra" in result:
             # Filter out less important mismatches if desired (e.g., attribute order)
             # For now, count all differences.
             significant_comparisons += 1


    if significant_comparisons == 0:
         # If only the root tag matched but nothing inside differed or was compared
         # Check if root tags did match (implicitly assumed if we got this far without root mismatch result)
         return 1.0 if comparison_results and ": Mismatch (Tag:" not in comparison_results[0] else 0.0

    # Ensure we don't divide by zero if no significant comparisons were counted
    return match_count / significant_comparisons if significant_comparisons > 0 else 0.0


def evaluate_part_score(similarity_score: float, total_points: float) -> float:
    """
    [R01, R04, R05]
    Calculates the final points for a part based on its similarity score and total possible points.
    Applies the formula: final_points = similarity_score * total_points.
    Clamps the final points to be not more than total_points.
    """
    calculated_final_points = similarity_score * total_points
    final_points = min(calculated_final_points, total_points) # R05: Clamping
    return round(final_points, 2) # Added rounding


def compare_xml_elements(ref_elem: ET.Element, student_elem: ET.Element, path: str = "") -> List[str]:
    """
    Recursively compares two XML elements and their children.
    Reports matches and mismatches based on element tags, attributes, and text.
    Handles 'Part' element children (features) as unordered.

    Args:
        ref_elem: The reference XML element.
        student_elem: The student XML element.
        path: The current hierarchical path for reporting.

    Returns:
        A list of strings, each representing a comparison result.
    """
    results = []

    current_path = f"{path}:{ref_elem.tag}" if path else ref_elem.tag

    # 1. Compare Element Tags
    if ref_elem.tag != student_elem.tag:
        results.append(f"{current_path} : Mismatch (Tag: Reference='{ref_elem.tag}', Student='{student_elem.tag}')")
        return results

    # 2. Compare Attributes
    ref_attrs = dict(ref_elem.items())
    student_attrs = dict(student_elem.items())
    all_attr_names = set(ref_attrs.keys()) | set(student_attrs.keys()) # Consider all unique attrs

    for attr_name in sorted(list(all_attr_names)): # Sort for consistent reporting
        ref_attr_value = ref_attrs.get(attr_name)
        student_attr_value = student_attrs.get(attr_name)

        if ref_attr_value is not None and student_attr_value is not None:
            normalized_ref = normalize_value(ref_attr_value)
            normalized_student = normalize_value(student_attr_value)

            if isinstance(normalized_ref, float) and isinstance(normalized_student, float):
                if abs(normalized_ref - normalized_student) < FLOAT_TOLERANCE:
                    results.append(f"{current_path} Attribute '{attr_name}' : Match (Value='{ref_attr_value}')")
                else:
                    results.append(f"{current_path} Attribute '{attr_name}' : Mismatch (Ref='{ref_attr_value}', Student='{student_attr_value}')")
            elif normalized_ref == normalized_student:
                results.append(f"{current_path} Attribute '{attr_name}' : Match (Value='{ref_attr_value}')")
            else:
                results.append(f"{current_path} Attribute '{attr_name}' : Mismatch (Ref='{ref_attr_value}', Student='{student_attr_value}')")
        elif ref_attr_value is not None:
            results.append(f"{current_path} Attribute '{attr_name}' : Missing in student model (Ref Value='{ref_attr_value}')")
        elif student_attr_value is not None:
            results.append(f"{current_path} Attribute '{attr_name}' : Extra in student model (Student Value='{student_attr_value}')")


    # 3. Compare Text Content
    ref_text = ref_elem.text.strip() if ref_elem.text else None # Use None if no text
    student_text = student_elem.text.strip() if student_elem.text else None

    if ref_text is not None and student_text is not None:
        normalized_ref_text = normalize_value(ref_text)
        normalized_student_text = normalize_value(student_text)

        if isinstance(normalized_ref_text, float) and isinstance(normalized_student_text, float):
            if abs(normalized_ref_text - normalized_student_text) < FLOAT_TOLERANCE:
                results.append(f"{current_path} Text : Match (Value='{ref_text}')")
            else:
                results.append(f"{current_path} Text : Mismatch (Ref='{ref_text}', Student='{student_text}')")
        elif normalized_ref_text == normalized_student_text:
            results.append(f"{current_path} Text : Match (Value='{ref_text}')")
        else:
             results.append(f"{current_path} Text : Mismatch (Ref='{ref_text}', Student='{student_text}')")
    elif ref_text is not None:
        results.append(f"{current_path} Text : Missing in student model (Ref Text='{ref_text}')")
    elif student_text is not None:
         results.append(f"{current_path} Text : Extra in student model (Student Text='{student_text}')")
    # If both are None, it's a match - no report line needed.


    # 4. Compare Children: Special handling for 'Part' element (features are unordered)
    if ref_elem.tag == 'Part':

        ref_physical_props = next((child for child in ref_elem if child.tag == 'PhysicalProperties'), None)
        student_physical_props = next((child for child in student_elem if child.tag == 'PhysicalProperties'), None)

        ref_features = [child for child in ref_elem if child.tag != 'PhysicalProperties']
        student_features_pool = [child for child in student_elem if child.tag != 'PhysicalProperties']

        # Keeping track of student features that have been matched
        matched_student_features_indices = set()

        # Compare PhysicalProperties first (treated as ordered/unique)
        if ref_physical_props is not None and student_physical_props is not None:
            results.extend(compare_xml_elements(
                ref_physical_props,
                student_physical_props,
                current_path # Pass the Part path
            ))
        elif ref_physical_props is not None:
            results.append(f"{current_path}:PhysicalProperties : Missing in student model")
        elif student_physical_props is not None:
            results.append(f"{current_path}:PhysicalProperties : Extra in student model")

        # Compare other features (unordered, using similarity)
        for ref_feature in ref_features:
            best_match_student_feature = None
            highest_similarity = -1.0
            best_match_index = -1

            for i, student_feature in enumerate(student_features_pool):
                if i not in matched_student_features_indices:
                    similarity = calculate_element_similarity(ref_feature, student_feature)
                    if similarity > highest_similarity:
                        highest_similarity = similarity
                        best_match_student_feature = student_feature
                        best_match_index = i

            ref_feature_identifier = ref_feature.get('Type', ref_feature.tag) # Prefer 'Type' attribute if available
            if best_match_student_feature is not None and highest_similarity >= FEATURE_MATCH_THRESHOLD:
                results.append(f"{current_path}:Feature '{ref_feature_identifier}' : Match found (Similarity {highest_similarity:.2f})")
                # Recursively compare the matched features for details
                results.extend(compare_xml_elements(
                    ref_feature,
                    best_match_student_feature,
                    f"{current_path}:{ref_feature_identifier}" # Pass feature identifier in path
                ))
                matched_student_features_indices.add(best_match_index)
            else:
                results.append(f"{current_path}:Feature '{ref_feature_identifier}' : Missing in student model (No similar feature found or below threshold)")

        # Report any remaining student features as extra
        for i, student_feature in enumerate(student_features_pool):
            if i not in matched_student_features_indices:
                student_feature_identifier = student_feature.get('Type', student_feature.tag)
                results.append(f"{current_path}:Feature '{student_feature_identifier}' : Extra in student model")

    else: # Normal ordered child comparison
        ref_children = list(ref_elem)
        student_children = list(student_elem)
        len_ref = len(ref_children)
        len_student = len(student_children)

        for i in range(max(len_ref, len_student)):
            ref_child = ref_children[i] if i < len_ref else None
            student_child = student_children[i] if i < len_student else None

            if ref_child is not None and student_child is not None:
                 # Recursive call for paired children
                 results.extend(compare_xml_elements(ref_child, student_child, current_path))
            elif ref_child is not None:
                results.append(f"{current_path}:{ref_child.tag} : Missing in student model at position {i}")
            elif student_child is not None:
                results.append(f"{current_path}:{student_child.tag} : Extra in student model at position {i}")

    return results


def compare_xml_files(reference_file_path: str, student_file_path: str) -> Tuple[List[str], float]:
    """
    Compares two XML files and returns a list of comparison results and a part-level similarity score.
    [R03]
    """
    try:
        ref_tree = ET.parse(reference_file_path)
        ref_root = ref_tree.getroot()
    except ET.ParseError as e:
        return [f"Error parsing reference XML '{reference_file_path}': {e}"], 0.0
    except FileNotFoundError:
        return [f"Reference XML file not found: '{reference_file_path}'"], 0.0

    try:
        student_tree = ET.parse(student_file_path)
        student_root = student_tree.getroot()
    except ET.ParseError as e:
        return [f"Error parsing student XML '{student_file_path}': {e}"], 0.0
    except FileNotFoundError:
        return [f"Student XML file not found: '{student_file_path}'"], 0.0

    comparison_results = compare_xml_elements(ref_root, student_root)
    part_similarity_score = calculate_part_similarity(comparison_results) # [Sr01, Sr03]

    return comparison_results, part_similarity_score

def generate_excel_report(student_scores: Dict[str, Dict[str, Tuple[float, float]]], output_folder: str):
    """
    Generates a single Excel (.xlsx) file with two sheets:
    1. Detailed Scores: Contains the score for each part.
    2. Total Scores: Contains the final summarized score for each student.
    """
    report_path = os.path.join(output_folder, 'student_scores_report.xlsx')
    print(f"\nGenerating Excel report with multiple sheets at: {report_path}") # Use stdout for info

    try:
        # --- Prepare data for Sheet 1: Detailed Scores ---
        detailed_data = []
        for matriculation_num, parts_scores in student_scores.items():
            for part_name, (final_pts, total_pts) in parts_scores.items():
                detailed_data.append({
                    'Matriculation_Number': matriculation_num,
                    'Part_Name': part_name,
                    'Final_Score': final_pts,
                    'Total_Possible_Points': total_pts
                })
        detailed_df = pd.DataFrame(detailed_data)

        # --- Prepare data for Sheet 2: Total Scores ---
        total_data = []
        for matriculation_num, parts_scores in student_scores.items():
            total_student_score = sum(score[0] for score in parts_scores.values())
            total_possible_score = sum(score[1] for score in parts_scores.values())
            total_data.append({
                'Matriculation_Number': matriculation_num,
                'Total_Final_Score': total_student_score,
                'Total_Possible_Points': total_possible_score
            })
        total_df = pd.DataFrame(total_data)

        detailed_df['Final_Score'] = detailed_df['Final_Score'].round(2)
        total_df['Total_Final_Score'] = total_df['Total_Final_Score'].round(2)

        # Ensure correct types - Use object/string for Matriculation Number if it might not be purely numeric
        detailed_df = detailed_df.astype({
            'Matriculation_Number': 'string', # Changed to string for safety
            'Part_Name': 'string',
            'Final_Score': 'float64',
            'Total_Possible_Points': 'float64' # Changed to float for consistency
        })

        total_df = total_df.astype({
            'Matriculation_Number': 'string', # Changed to string
            'Total_Final_Score': 'float64',
            'Total_Possible_Points': 'float64' # Changed to float
        })

        # --- Write both DataFrames to a single Excel file on different sheets ---
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            detailed_df.to_excel(writer, sheet_name='Detailed Scores', index=False)
            total_df.to_excel(writer, sheet_name='Total Scores', index=False)

        print("Excel report generated successfully.") # Use stdout for info
    except Exception as e:
         # Use stderr for errors
        print(f"Error: Could not generate Excel report. {e}", file=sys.stderr)
        # Re-raise the exception so the script exits with an error code
        raise e

# --- MODIFICATION: To add a new function that encapsulates the main logic for comparison, allowing it to be called from C# with arguments instead of relying on input().---

def run_comparison_logic(args):
    """
    Contains the core logic from your original __main__ block,
    but uses parsed arguments instead of input().
    This function WILL NOT BE MODIFIED from your baseline logic,
    except to replace input() calls with args properties.
    """
    ref_xml_folder = args.baseline # Replaces input()
    student_xml_folder = args.input # Replaces input()
    output_folder_path = args.output # Replaces input()
    points_map_json = args.points_map # Comes from C#

    # --- Load Points Map (Modified from input() to use args.points_map) ---
    try:
        # Handle potential extra quotes if passed incorrectly from cmd line
        if points_map_json.startswith('"') and points_map_json.endswith('"'):
            points_map_json = points_map_json[1:-1].replace('\\"', '"')
        reference_part_total_points = json.loads(points_map_json)
        # Convert points to float for calculations
        for part, points in reference_part_total_points.items():
            try:
                reference_part_total_points[part] = float(points)
            except ValueError:
                # Use stderr for errors
                print(f"Error: Invalid points value '{points}' for part '{part}' in points map. Must be numeric.", file=sys.stderr)
                sys.exit(1) # Exit with failure
        # Use stdout for informational messages
        print(f"Loaded points map: {reference_part_total_points}")
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON string provided for --points-map: {points_map_json}. Details: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error loading points map: {e}", file=sys.stderr)
        sys.exit(1)


    # --- Load Reference Models (Error handling added, uses args.baseline) ---
    reference_models: List[Tuple[str, str, ET.Element]] = []
    ref_files_in_folder = [f for f in os.listdir(ref_xml_folder) if f.lower().endswith('.xml')]

    if not ref_files_in_folder:
        print(f"Error: No XML files found in the reference folder: '{ref_xml_folder}'. Exiting.", file=sys.stderr)
        sys.exit(1)

    print(f"\nLoading {len(ref_files_in_folder)} reference XML files...") # Info to stdout
    for ref_file_name in ref_files_in_folder:
        ref_file_path = os.path.join(ref_xml_folder, ref_file_name)
        ref_filename_base = os.path.splitext(ref_file_name)[0]
        try:
            ref_tree = ET.parse(ref_file_path)
            reference_models.append((ref_filename_base, ref_file_path, ref_tree.getroot()))
            print(f"  Loaded: {ref_file_name}") # Info to stdout
        except ET.ParseError as e:
             # Use stderr for warnings
            print(f"  Warning: Error parsing reference XML '{ref_file_name}': {e}. Skipping.", file=sys.stderr)
        except FileNotFoundError:
            print(f"  Warning: Reference XML file not found: '{ref_file_name}'. Skipping.", file=sys.stderr)

    if not reference_models:
        print("Error: No valid reference XML files were loaded. Exiting.", file=sys.stderr)
        sys.exit(1)

    # --- Create Output Folder (Error handling added, uses args.output) ---
    try:
        os.makedirs(output_folder_path, exist_ok=True)
    except OSError as e:
        print(f"Error: Could not create output directory '{output_folder_path}'. {e}", file=sys.stderr)
        print("Please check permissions and path validity. Exiting.", file=sys.stderr)
        sys.exit(1)


    # --- Process Student Files (Uses args.input, logic unchanged) ---
    student_files = [f for f in os.listdir(student_xml_folder) if f.lower().endswith('.xml')]
    student_scores: Dict[str, Dict[str, Tuple[float, float]]] = {} # {matriculation_num: {part_name: (final_points, total_points)}}

    if not student_files:
        print(f"\nWarning: No XML files found in student folder '{student_xml_folder}'. No comparisons to perform.") # Info stdout
    else:
        print(f"\nFound {len(student_files)} student XML files to compare.") # Info stdout
        for student_file_name in student_files:
            student_file_path = os.path.join(student_xml_folder, student_file_name)
            student_filename_base = os.path.splitext(student_file_name)[0]

            # Extract student matriculation number and part name (robustly)
            parts = student_filename_base.split('_', 1)
            student_matriculation_num = parts[0] if len(parts) > 1 else "UNKNOWN"
            student_part_name = parts[1] if len(parts) > 1 else student_filename_base

            if student_matriculation_num not in student_scores:
                student_scores[student_matriculation_num] = {}

            best_ref_match_info = None
            # Find the best matching reference file based on part name (case-insensitive)
            for ref_base_name, ref_full_path, ref_root_elem in reference_models:
                ref_part_name = ref_base_name.split('_', 1)[-1] # Get part name portion

                if ref_part_name.lower() == student_part_name.lower():
                    best_ref_match_info = (ref_base_name, ref_full_path, ref_root_elem, ref_part_name)
                    break # Found exact match

            if best_ref_match_info is None:
                # Info to stdout, warning to stderr
                print(f"\n--- Skipping Student File: {student_file_name} ---")
                print(f"  Warning: No suitable reference file found for part name '{student_part_name}'.", file=sys.stderr)
                continue # Skip to the next student file

            matched_ref_filename_base, matched_ref_full_path, matched_ref_root_element, matched_ref_part_name = best_ref_match_info

            print(f"\n--- Comparing Student File: {student_file_name} ---") # Info stdout
            print(f"  Matched with Reference: {os.path.basename(matched_ref_full_path)}") # Info stdout

            # Perform comparison using your original functions
            results, part_similarity_score = compare_xml_files(matched_ref_full_path, student_file_path)

            # Print detailed comparison results (Uncomment if needed, prints to stdout)
            # for line in results:
            #    print(line)

            # Get total points, handling case-insensitivity and missing points
            total_points_for_part = 0.0
            found_points = False
            for map_part_name, map_points_val in reference_part_total_points.items():
                if map_part_name.lower() == matched_ref_part_name.lower():
                    total_points_for_part = float(map_points_val) # Already converted
                    found_points = True
                    break
            if not found_points:
                 # Warning to stderr
                print(f"  Warning: No points defined in points map for reference part '{matched_ref_part_name}'. Using 0 points.", file=sys.stderr)

            # Evaluate score using your original function
            final_points_for_part = evaluate_part_score(part_similarity_score, total_points_for_part)

            # Store score
            student_scores[student_matriculation_num][student_part_name] = (final_points_for_part, total_points_for_part)

            # Print summary for this part to stdout
            print(f"  Part Similarity Score: {part_similarity_score:.2f}")
            print(f"  Final Score for '{student_part_name}': {final_points_for_part:.2f} out of {total_points_for_part:.2f}")
            print(f"--- Finished Comparison for {student_file_name} ---")

    # --- Print Summary (Logic unchanged, uses stdout) ---
    print("\n--- Summary of Student Scores ---")
    if not student_scores:
        print("No student scores to report.")
    else:
        # Sort by matriculation number (attempt numeric sort first)
        try:
             sorted_matriculation_nums = sorted(student_scores.keys(), key=int)
        except ValueError:
             sorted_matriculation_nums = sorted(student_scores.keys()) # Fallback to string sort

        for matriculation_num in sorted_matriculation_nums:
            parts_scores = student_scores[matriculation_num]
            print(f"\nStudent: {matriculation_num}")
            total_student_score = 0.0
            total_possible_score = 0.0
            for part_name in sorted(parts_scores.keys()): # Sort parts alphabetically
                final_pts, total_pts = parts_scores[part_name]
                print(f"    - Part '{part_name}': Score - {final_pts:.2f} out of {total_pts:.2f}")
                total_student_score += final_pts
                total_possible_score += total_pts
            if total_possible_score > 0:
                print(f"    -> Total Student Score: {total_student_score:.2f} out of {total_possible_score:.2f}")

    # --- Generate Report (Logic unchanged, uses args.output) ---
    if student_scores:
        try:
            generate_excel_report(student_scores, output_folder_path)
        except Exception as e:
             # Report error but allow script to finish if desired
             print(f"\nError occurred during Excel report generation: {e}", file=sys.stderr)
             # sys.exit(1) # Optionally exit with error if report generation is critical


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="XML CAD Model Comparison Tool - Integrated Version")
    # Arguments for full comparison mode
    parser.add_argument("--input", help="Student XML folder path (required for full comparison).")
    parser.add_argument("--baseline", required=True, help="Reference XML folder path (required for both modes).")
    parser.add_argument("--output", help="Output report folder path (required for full comparison).")
    parser.add_argument("--points-map", help="JSON string mapping part names to points (required for full comparison).")
    # Argument to switch to list-parts mode
    parser.add_argument("--list-parts", action="store_true", help="If set, lists unique part names from baseline folder and exits.")

    args = parser.parse_args()

    try:
        # --- MODE 1: List Parts ---
        if args.list_parts:
            # Check baseline folder exists
            if not os.path.isdir(args.baseline):
                print(f"Error: Baseline folder not found: {args.baseline}", file=sys.stderr)
                sys.exit(1)

            ref_files = [f for f in os.listdir(args.baseline) if f.lower().endswith('.xml')]
            if not ref_files:
                print(f"Error: No XML files found in baseline folder '{args.baseline}'.", file=sys.stderr)
                sys.exit(1)

            part_names = set()
            for ref_file_name in ref_files:
                ref_filename_base = os.path.splitext(ref_file_name)[0]
                # Extract part name (last part after first '_', or full name if no '_')
                part_name = ref_filename_base.split('_', 1)[-1]
                if part_name not in part_names:
                     # --- Print ONLY the part name to stdout for C# to capture ---
                    print(part_name)
                    part_names.add(part_name)
            sys.exit(0) # Exit successfully after listing parts

        # --- MODE 2: Full Comparison ---
        else:
            # Check required arguments for full mode
            required_args_full = [args.input, args.output, args.points_map]
            if not all(required_args_full):
                 # Improved error message listing missing arguments
                 missing = []
                 if not args.input: missing.append("--input")
                 if not args.output: missing.append("--output")
                 if not args.points_map: missing.append("--points-map")
                 print(f"Error: Missing required arguments for full comparison: {', '.join(missing)}", file=sys.stderr)
                 parser.print_help(file=sys.stderr) # Show help message
                 sys.exit(1)

            # Validate input/baseline folders exist
            if not os.path.isdir(args.input):
                print(f"Error: Student input folder not found: {args.input}", file=sys.stderr)
                sys.exit(1)
            # Baseline folder existence is already checked if --list-parts is not used, but double-check here
            if not os.path.isdir(args.baseline):
                 print(f"Error: Baseline folder not found: {args.baseline}", file=sys.stderr)
                 sys.exit(1)

            # Call the main comparison logic function using the parsed arguments
            run_comparison_logic(args)
            print("\nComparison process completed successfully.") # Final success message to stdout
            sys.exit(0) # Explicit success exit

    except Exception as e:
        # Catch any unexpected error during execution
        print(f"\nFATAL ERROR: An unexpected error occurred: {e}", file=sys.stderr)
        import traceback
        print("\n--- Traceback ---", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        print("--- End Traceback ---", file=sys.stderr)
        sys.exit(1) # Exit with failure code

