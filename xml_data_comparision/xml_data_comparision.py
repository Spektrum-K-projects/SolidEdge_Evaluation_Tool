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

    children1 = list(elem1)
    children2 = list(elem2)
    
    max_score += len(children1) * 10 

    child_tags1 = [c.tag for c in children1]
    child_tags2 = [c.tag for c in children2]

    for tag in set(child_tags1):
        count1 = child_tags1.count(tag)
        count2 = child_tags2.count(tag)
        score += min(count1, count2) * 10 

    if max_score == 0:
        return 1.0
    
    return score / max_score

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

def compare_xml_files(reference_file_path: str, student_file_path: str) -> List[str]:
    """
    Compares two XML files and returns a list of comparison results.
    """
    try:
        ref_tree = ET.parse(reference_file_path)
        ref_root = ref_tree.getroot()
    except ET.ParseError as e:
        return [f"Error parsing reference XML '{reference_file_path}': {e}"]
    except FileNotFoundError:
        return [f"Reference XML file not found: '{reference_file_path}'"]

    try:
        student_tree = ET.parse(student_file_path)
        student_root = student_tree.getroot()
    except ET.ParseError as e:
        return [f"Error parsing student XML '{student_file_path}': {e}"]
    except FileNotFoundError:
        return [f"Student XML file not found: '{student_file_path}'"]

    return compare_xml_elements(ref_root, student_root)

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

    student_xml_folder = input("Enter the full path to the folder containing student XML files: ").strip()
    while not os.path.isdir(student_xml_folder):                        #Error Handling
        print("Error: Student XML folder not found. Please try again.")
        student_xml_folder = input("Enter the full path to the folder containing student XML files: ").strip()

    student_files = [f for f in os.listdir(student_xml_folder) if f.lower().endswith('.xml')]

    if not student_files:
        print(f"No XML files found in '{student_xml_folder}'.")
    else:
        print(f"\nFound {len(student_files)} student XML files to compare.")
        for student_file_name in student_files:
            student_file_path = os.path.join(student_xml_folder, student_file_name)
            student_filename_base = os.path.splitext(student_file_name)[0]

            best_ref_match_info = None                                   # (ref_filename_base, ref_file_path, ref_root_element)
            max_similarity_found = -1.0

            # Find the best matching reference file for the current student file
            for ref_base_name, ref_full_path, ref_root_elem in reference_models:
                similarity_ratio = calculate_filename_similarity(ref_base_name, student_filename_base)
                
                if similarity_ratio > max_similarity_found and similarity_ratio >= FILENAME_SIMILARITY_THRESHOLD:
                    max_similarity_found = similarity_ratio
                    best_ref_match_info = (ref_base_name, ref_full_path, ref_root_elem)
            
            if best_ref_match_info is None:
                print(f"\n--- Skipping Student File: {student_file_name} ---")
                print(f"No suitable reference file found for '{student_filename_base}' (max similarity {max_similarity_found:.2f} < {FILENAME_SIMILARITY_THRESHOLD:.2f})")
                print(f"--- Skipped Comparison for {student_file_name} ---")
                continue                                                    # Skip to the next student file

            matched_ref_filename_base, matched_ref_full_path, matched_ref_root_element = best_ref_match_info

            print(f"\n--- Comparing Student File: {student_file_name} ---")
            print(f"  Matched with Reference: {os.path.basename(matched_ref_full_path)} (Filename Similarity: {max_similarity_found:.2f})")
            
            results = compare_xml_files(matched_ref_full_path, student_file_path)
            
            for line in results:
                print(line)
            print(f"--- Finished Comparison for {student_file_name} ---")

    print("\nComparison process completed.")