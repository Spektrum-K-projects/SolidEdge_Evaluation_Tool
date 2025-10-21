import os
import xml.etree.ElementTree as ET
import difflib # For string similarity comparison
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
            matched_children_indices.add(best_child_match_index)
    
    max_score += len(children1) * 10 # Add potential score for all children

    if max_score == 0:
        return 1.0 # Both elements are empty, consider them a perfect match
    
    return score / max_score

def calculate_part_similarity(comparison_results: List[str]) -> float:
    """
    [Sr01, Sr03]
    Calculates a single part-level similarity score based on the detailed comparison results.
    This implementation counts 'Match' results vs. total results for simplicity.
    Can be expanded for more sophisticated metrics (e.g., weighting different types of matches/mismatches).
    """
    total_results = len(comparison_results)
    if total_results == 0:
        return 1.0 # No comparisons, assume perfect similarity

    match_count = 0
    # A simplified approach to count matches.
    # A more robust approach would categorize results (Match, Mismatch, Missing, Extra)
    # and assign weights to each category.
    for result in comparison_results:
        if "Match" in result and "Mismatch" not in result and "Missing" not in result and "Extra" not in result:
            match_count += 1
    
    # This is a very basic similarity. A more advanced one might involve:
    # 1. Counting specific "feature matched" lines vs. total features.
    # 2. Assigning weights to different types of matches (e.g., tag match > attribute match > text match).
    # 3. Penalizing missing/extra elements more heavily.
    return match_count / total_results

def evaluate_part_score(similarity_score: float, total_points: float) -> float:
    """
    [R01, R04, R05]
    Calculates the final points for a part based on its similarity score and total possible points.
    Applies the formula: final_points = similarity_score * total_points.
    Clamps the final points to be not more than total_points.
    """
    calculated_final_points = similarity_score * total_points
    final_points = min(calculated_final_points, total_points) # R05: Clamping
    return final_points

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

    for attr_name, ref_attr_value in ref_attrs.items():
        if attr_name in student_attrs:
            student_attr_value = student_attrs[attr_name]
            
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
        else:
            results.append(f"{current_path} Attribute '{attr_name}' : Missing in student model")
    
    # Check for extra attributes in student model
    for attr_name in student_attrs:
        if attr_name not in ref_attrs:
            results.append(f"{current_path} Attribute '{attr_name}' : Extra in student model")

    # 3. Compare Text Content
    ref_text = ref_elem.text.strip() if ref_elem.text else "NO VALUE"
    student_text = student_elem.text.strip() if student_elem.text else "NO VALUE"

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

    # 4. Compare Children: Special handling for 'Part' element (features are unordered)
    if ref_elem.tag == 'Part':

        ref_physical_props = next((child for child in ref_elem if child.tag == 'PhysicalProperties'), None)
        student_physical_props = next((child for child in student_elem if child.tag == 'PhysicalProperties'), None)

        ref_features = [child for child in ref_elem if child.tag != 'PhysicalProperties']
        student_features_pool = [child for child in student_elem if child.tag != 'PhysicalProperties']
        
        # Keeping track of student features that have been matched
        matched_student_features_indices = set()

        # Compare PhysicalProperties first
        if ref_physical_props is not None and student_physical_props is not None:
            results.extend(compare_xml_elements(
                ref_physical_props,
                student_physical_props,
                current_path
            ))
        elif ref_physical_props is not None:
            results.append(f"{current_path}:PhysicalProperties : Missing in student model")
        elif student_physical_props is not None:
            results.append(f"{current_path}:PhysicalProperties : Extra in student model")

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
            
            if best_match_student_feature is not None and highest_similarity >= FEATURE_MATCH_THRESHOLD:
                results.append(f"{current_path}:{ref_feature.tag} (Matched by similarity {highest_similarity:.2f}) : Match")
                results.extend(compare_xml_elements(
                    ref_feature,
                    best_match_student_feature,
                    current_path
                ))
                matched_student_features_indices.add(best_match_index)
            else:
                results.append(f"{current_path}:{ref_feature.tag} : Missing in student model (No similar feature found or below {FEATURE_MATCH_THRESHOLD*100}% similarity)")

        for i, student_feature in enumerate(student_features_pool):
            if i not in matched_student_features_indices:
                results.append(f"{current_path}:{student_feature.tag} : Extra in student model")

    else:
        ref_children = list(ref_elem)
        student_children = list(student_elem)

        for i in range(max(len(ref_children), len(student_children))):
            ref_child = ref_children[i] if i < len(ref_children) else None
            student_child = student_children[i] if i < len(student_children) else None

            if ref_child is not None and student_child is not None:
                results.extend(compare_xml_elements(ref_child, student_child, current_path))
            elif ref_child is not None:
                results.append(f"{current_path}:{ref_child.tag} : Missing in student model")
            elif student_child is not None:
                results.append(f"{current_path}:{student_child.tag} : Extra in student model")

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

if __name__ == "__main__":
    print("--- XML CAD Model Comparison Tool ---")

    ref_xml_folder = input("Enter the full path to the folder containing reference XML files: ").strip()
    while not os.path.isdir(ref_xml_folder):
        print("Error: Reference XML folder not found. Please try again.")
        ref_xml_folder = input("Enter the full path to the folder containing reference XML files: ").strip()

    reference_models: List[Tuple[str, str, ET.Element]] = [] # (base_filename, full_path, root_element)
    ref_files_in_folder = [f for f in os.listdir(ref_xml_folder) if f.lower().endswith('.xml')]

    if not ref_files_in_folder:
        print(f"No XML files found in the reference folder: '{ref_xml_folder}'. Exiting.")
        exit()

    print(f"\nLoading {len(ref_files_in_folder)} reference XML files...")
    for ref_file_name in ref_files_in_folder:
        ref_file_path = os.path.join(ref_xml_folder, ref_file_name)
        ref_filename_base = os.path.splitext(ref_file_name)[0]
        try:
            ref_tree = ET.parse(ref_file_path)
            reference_models.append((ref_filename_base, ref_file_path, ref_tree.getroot()))
            print(f"  Loaded: {ref_file_name}")
        except ET.ParseError as e:
            print(f"  Error parsing reference XML '{ref_file_name}': {e}. Skipping.")
        except FileNotFoundError:
            print(f"  Reference XML file not found: '{ref_file_name}'. Skipping.")

    if not reference_models:
        print("No valid reference XML files were loaded. Exiting.")
        exit()

    # [R02] User input for total points per reference part
    reference_part_total_points: Dict[str, float] = {}
    print("\nPlease enter the total points for each reference part:")
    for ref_filename_base, _, _ in reference_models:
        while True:
            try:
                # Extract part name (e.g., 'part_a' from '123456_part_a.xml')
                # Assuming filename convention: [MatriculationNumber]_[PartName].xml
                # If the ref filename does not have matriculation number, it's just the part name.
                part_name_for_prompt = ref_filename_base.split('_', 1)[1] if '_' in ref_filename_base else ref_filename_base
                points_input = input(f"  Enter total points for '{part_name_for_prompt}': ").strip()
                total_pts = float(points_input)
                if total_pts < 0:
                    print("Total points cannot be negative. Please enter a non-negative number.")
                else:
                    reference_part_total_points[part_name_for_prompt] = total_pts
                    break
            except ValueError:
                print("Invalid input. Please enter a numeric value for total points.")

    student_xml_folder = input("\nEnter the full path to the folder containing student XML files: ").strip()
    while not os.path.isdir(student_xml_folder):                        #Error Handling
        print("Error: Student XML folder not found. Please try again.")
        student_xml_folder = input("Enter the full path to the folder containing student XML files: ").strip()

    student_files = [f for f in os.listdir(student_xml_folder) if f.lower().endswith('.xml')]

    # [R06] Data structure to store student scores
    student_scores: Dict[str, Dict[str, Tuple[float, float]]] = {} # {matriculation_num: {part_name: (final_points, total_points)}}

    if not student_files:
        print(f"No XML files found in '{student_xml_folder}'.")
    else:
        print(f"\nFound {len(student_files)} student XML files to compare.")
        for student_file_name in student_files:
            student_file_path = os.path.join(student_xml_folder, student_file_name)
            student_filename_base = os.path.splitext(student_file_name)[0]

            # Extract student matriculation number and part name
            parts = student_filename_base.split('_', 1)
            if len(parts) >= 2:
                student_matriculation_num = parts[0]
                student_part_name = parts[1]
            else:
                student_matriculation_num = "UNKNOWN_STUDENT"
                student_part_name = student_filename_base
            
            # Initialize student's score entry if not present
            if student_matriculation_num not in student_scores:
                student_scores[student_matriculation_num] = {}

            best_ref_match_info = None                                   # (ref_filename_base, ref_full_path, ref_root_element)
            max_similarity_found = -1.0

            # Find the best matching reference file for the current student file
            # This logic uses the base filename (e.g., "part_a") for matching, not the matriculation number
            for ref_base_name, ref_full_path, ref_root_elem in reference_models:
                # Extract part name from reference file for comparison (e.g., 'part_a')
                ref_part_name = ref_base_name.split('_', 1)[1] if '_' in ref_base_name else ref_base_name
                
                # Check if the student's part name matches a reference part name
                # This ensures we are comparing 'part_a' student file with 'part_a' reference file
                if ref_part_name.lower() == student_part_name.lower():
                    # Calculate similarity based on the *full* base names for the purpose of the initial filename similarity check
                    # However, the primary match is on the extracted part name.
                    similarity_ratio = calculate_filename_similarity(ref_base_name, student_filename_base)
                    
                    if similarity_ratio > max_similarity_found: # No FILENAME_SIMILARITY_THRESHOLD here as we are doing exact part name match
                        max_similarity_found = similarity_ratio
                        best_ref_match_info = (ref_base_name, ref_full_path, ref_root_elem, ref_part_name)
                    # We found an exact part name match, so we don't need to look further for this student part
                    break 
            
            if best_ref_match_info is None:
                print(f"\n--- Skipping Student File: {student_file_name} ---")
                print(f"No suitable reference file found for '{student_part_name}' (no matching reference part name)")
                print(f"--- Skipped Comparison for {student_file_name} ---")
                continue                                                    # Skip to the next student file

            matched_ref_filename_base, matched_ref_full_path, matched_ref_root_element, matched_ref_part_name = best_ref_match_info

            print(f"\n--- Comparing Student File: {student_file_name} ---")
            print(f"  Matched with Reference: {os.path.basename(matched_ref_full_path)} (Filename Similarity: {max_similarity_found:.2f})")
            
            # [R03] Get comparison results and part-level similarity score
            results, part_similarity_score = compare_xml_files(matched_ref_full_path, student_file_path)
            
            for line in results:
                print(line)

            # [R07] Get total points for this part
            total_points_for_part = reference_part_total_points.get(matched_ref_part_name, 0.0)
            
            # [R01, R04, R05] Evaluate the score for the current part
            final_points_for_part = evaluate_part_score(part_similarity_score, total_points_for_part)

            # [Sr02] Store the part-level similarity score (implicitly, as part of final_points calc)
            # [R06] Store the student's score for this part
            student_scores[student_matriculation_num][student_part_name] = (final_points_for_part, total_points_for_part)

            print(f"  Part Similarity Score: {part_similarity_score:.2f}")
            print(f"  Final Score for '{student_part_name}': {final_points_for_part:.2f} out of {total_points_for_part:.2f}")
            print(f"--- Finished Comparison for {student_file_name} ---")

    print("\n--- Summary of Student Scores ---")
    if not student_scores:
        print("No student scores to report.")
    else:
        for matriculation_num, parts_scores in student_scores.items():
            print(f"{matriculation_num}: Student Score")
            for part_name, (final_pts, total_pts) in parts_scores.items():
                print(f"    - Part {part_name.replace('part_', '').upper()}: Score - {final_pts:.2f} out of {total_pts:.2f}")

    print("\nComparison process completed.")