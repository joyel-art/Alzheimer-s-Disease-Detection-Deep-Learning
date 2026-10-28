import os
import numpy as np
import tensorflow as tf
from PIL import Image
import tf_keras
import warnings

# Hardcoded classes based on dataset directory structure
# ImageDataGenerator sorts directories alphabetically
CLASSES = ['MildDemented', 'ModerateDemented', 'NonDemented', 'VeryMildDemented']

_model = None

def get_model():
    global _model
    if _model is None:
        # Resolve absolute path to the model file
        current_dir = os.path.dirname(os.path.abspath(__file__))
        model_path = os.path.join(current_dir, 'code', 'model.h5')
        
        if not os.path.exists(model_path):
            # Try fallback to 'model.h5' in root if 'code/model.h5' fails
            model_path = os.path.join(current_dir, 'model.h5')
            if not os.path.exists(model_path):
                raise FileNotFoundError(f"Model file not found at standard locations.")
            
        print(f"Loading MRI diagnostic model from {model_path}...")
        # Use tf_keras for compatibility with the saved model optimizer (Adamax)
        _model = tf_keras.models.load_model(model_path)
        _model.compile(tf_keras.optimizers.Adamax(learning_rate=0.001), 
                       loss='categorical_crossentropy', metrics=['accuracy'])
    return _model

def is_brain_scan(image):
    """
    Heuristic check to ensure the image is likely a brain MRI scan.
    Returns (bool, reason_string)
    """
    try:
        # 1. Convert to RGB and numpy array
        img_rgb = image.convert('RGB')
        img_array = np.array(img_rgb)
        
        # 2. Check for "grayscale-ness" (typical of MRIs)
        # MRIs usually have very low color saturation
        r, g, b = img_array[:,:,0], img_array[:,:,1], img_array[:,:,2]
        color_variance = np.mean(np.abs(r.astype(int) - g.astype(int)) + np.abs(g.astype(int) - b.astype(int)))
        
        # 3. Check for dark background
        # MRIs usually have a large black area around the brain or dark tissue
        background_pixels = np.sum(img_array < 50) / img_array.size
        
        # 4. Check for edge density (MRI scans have soft tissue edges)
        # Simplified: Check the center of the image has some content vs edges
        h, w, _ = img_array.shape
        center_region = img_array[h//4:3*h//4, w//4:3*w//4]
        center_mean = np.mean(center_region)
        
        # Debug info
        print(f"DEBUG Analysis - Color Var: {color_variance:.2f}, Dark BG: {background_pixels:.4f}, Center Mean: {center_mean:.2f}")
        
        # Heuristic Thresholds:
        if color_variance > 5: 
            return False, "Input image must be a black and white (grayscale) MRI scan. Colored images are not allowed."
        if background_pixels < 0.10:
            return False, "Image lacks the characteristic background distribution of an MRI."
        if center_mean < 20:
             return False, "Image appears mostly empty or too dark."
            
        return True, "Valid"
    except Exception:
        return False, "Internal error during image validation."

def alzhim(inpath):
    """
    Main entry point for Deep Learning detection of Alzheimer stages.
    """
    warnings.filterwarnings("ignore")
    print(f"\n--- Starting Deep Learning MRI Analysis ---")
    print(f"Input Path: {inpath}")
    
    try:
        if not os.path.exists(inpath):
            return "Error: MRI file path not found on server."
            
        image = Image.open(inpath).convert("RGB")
        
        # 1. MRI Validation Check
        is_valid, val_reason = is_brain_scan(image)
        if not is_valid:
            print(f"Validation Failed: {val_reason}")
            return f"Invalid Scan: {val_reason}"
            
        # 2. Deep Learning Preprocessing
        # Model expects 224x224 RGB images
        img = image.resize((224, 224))
        img_array = tf.keras.preprocessing.image.img_to_array(img)
        img_array = np.expand_dims(img_array, axis=0)
        
        # 3. Model Prediction
        model = get_model()
        predictions = model.predict(img_array, verbose=0)
        
        # 4. Result Interpretation
        scores = predictions[0]
        max_idx = np.argmax(scores)
        confidence = scores[max_idx] * 100
        
        # If the model is extremely unconfident, it might be a malformed image
        if confidence < 30:
            return "Low Confidence: Scan quality/type unrecognizable by current model."
            
        predicted_class = CLASSES[max_idx]
        
        final_result = f"{predicted_class} ({confidence:.2f}%)"
        print(f"DEEP LEARNING RESULT: {final_result}")
        print(f"--- Analysis Complete ---\n")
        
        return final_result
        
    except Exception as e:
        print(f"CRITICAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"Analysis Error: {str(e)}"
