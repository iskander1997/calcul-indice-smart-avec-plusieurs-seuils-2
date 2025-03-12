import os
import pandas as pd
import re

# Detect if running in Colab
try:
    from google.colab import files
    import ipywidgets as widgets
    from IPython.display import display, HTML
    IN_COLAB = True
except ModuleNotFoundError:
    # For local Python environment, use tkinter
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    IN_COLAB = False

class ExcelFileSelector:
    def __init__(self, master=None, callback=None):
        # Store callback function
        self.callback = callback
        self.validation_result = None
        
        if IN_COLAB:
            # Variables for Colab implementation
            self.scores_path = None
            self.prices_path = None
            self.uploaded_files = {}
        else:
            # Variables for tkinter implementation
            self.root = tk.Tk() if master is None else tk.Toplevel(master)
            self.root.title("Sélection des Fichiers Excel")
            self.root.geometry("800x600")
            
            # Palette de couleurs moderne
            self.colors = {
                'background': 'ghost white',
                'primary': 'blue4',
                'secondary': 'blue4',
                'accent': '#2c3e50',
                'text': 'gray16',
                'white': '#ffffff'
            }
            
            # Configuration du style
            self.style = ttk.Style()
            self.style.theme_use('clam')
            
            # Configuration de la fenêtre
            self.root.configure(bg=self.colors['background'])
            self.root.resizable(False, False)
            
            # Variables de stockage
            self.prices_path = tk.StringVar()
            self.scores_path = tk.StringVar()
            self.thresholds = tk.StringVar()
        
    def create_ui(self):
        if IN_COLAB:
            # Create Colab UI with ipywidgets
            display(HTML("<h2 style='color: #2c3e50; text-align: center; margin-bottom: 20px;'>Sélection des Fichiers Excel</h2>"))
            
            # Create file upload widget for scores
            self.scores_upload = widgets.FileUpload(
                description='Scores des Sociétés:',
                accept='.xlsx, .xls',
                multiple=False,
                layout=widgets.Layout(width='auto', margin='10px 0')
            )
            self.scores_upload.observe(self.handle_scores_upload, names='value')
            
            # Create threshold input widget
            self.thresholds_input = widgets.Text(
                description='Seuils:',
                placeholder='ex: 10, 20, 30 (valeurs entre 1 et 199)',
                layout=widgets.Layout(width='auto', margin='10px 0')
            )
            
            # Create validation button
            self.validate_button = widgets.Button(
                description='Valider',
                button_style='primary',
                layout=widgets.Layout(width='auto', margin='20px 0 10px 0')
            )
            self.validate_button.on_click(self.handle_validate_button)
            
            # Status output area
            self.output = widgets.Output()
            
            # Display all widgets
            display(self.scores_upload)
            display(self.thresholds_input)
            display(self.validate_button)
            display(self.output)
        else:
            # Create tkinter UI
            # Cadre principal
            main_frame = tk.Frame(self.root, bg=self.colors['background'], padx=40, pady=30)
            main_frame.pack(expand=True, fill=tk.BOTH)
            
            # Titre
            title_label = tk.Label(main_frame, text="Sélection des Fichiers Excel", 
                                    font=('Segoe UI', 20, 'bold'), 
                                    fg=self.colors['accent'], 
                                    bg=self.colors['background'])
            title_label.pack(pady=(0, 30))
    
            # Fichier des Scores
            self.create_file_selector(main_frame, 
                                    "Scores des Sociétés", 
                                    self.scores_path, 
                                    self.select_scores_file)
            
            # Section Seuils
            self.create_thresholds_section(main_frame)
            
            # Bouton de Validation
            validate_button = tk.Button(main_frame, 
                                        text="Valider", 
                                        command=self.validate_inputs,
                                        bg=self.colors['primary'], 
                                        fg=self.colors['white'],
                                        font=('Segoe UI', 12, 'bold'),
                                        relief=tk.FLAT,
                                        activebackground=self.colors['secondary'])
            validate_button.pack(pady=(30, 0), ipadx=20, ipady=10)
    
    def handle_scores_upload(self, change):
        if not IN_COLAB:
            return
            
        if change['new']:
            # Get uploaded file name and content
            filename = next(iter(change['new']))
            file_content = change['new'][filename]['content']
            
            # Save the file content for later use
            self.uploaded_files[filename] = file_content
            
            # Set the paths
            self.scores_path = filename
            self.prices_path = filename
            
            with self.output:
                self.output.clear_output()
                print(f"Fichier sélectionné: {filename}")
    
    def handle_validate_button(self, button):
        """Handler for the validate button in Colab UI"""
        if not IN_COLAB:
            return
            
        self.validate_inputs()
    
    def create_file_selector(self, parent, label_text, path_var, select_command):
        if IN_COLAB:
            return
            
        # Cadre pour chaque sélecteur de fichier
        frame = tk.Frame(parent, bg=self.colors['background'])
        frame.pack(fill=tk.X, pady=10)
        
        # Label
        label = tk.Label(frame, 
                        text=label_text, 
                        font=('Segoe UI', 12), 
                        fg=self.colors['text'], 
                        bg=self.colors['background'])
        label.pack(anchor='w')
        
        # Sous-cadre pour l'entrée et les boutons
        input_frame = tk.Frame(frame, bg=self.colors['background'])
        input_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Entrée de fichier
        entry = tk.Entry(input_frame, 
                        textvariable=path_var, 
                        font=('Segoe UI', 10), 
                        width=60, 
                        state='readonly',
                        relief=tk.FLAT,
                        bg=self.colors['white'])
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        # Bouton Parcourir
        browse_btn = tk.Button(input_frame, 
                            text="Parcourir", 
                            command=select_command,
                            bg=self.colors['secondary'], 
                            fg=self.colors['white'],
                            font=('Segoe UI', 10),
                            relief=tk.FLAT,
                            activebackground=self.colors['primary'])
        browse_btn.pack(side=tk.LEFT)
        
        # Bouton Effacer
        clear_btn = tk.Button(input_frame, 
                            text="×", 
                            command=lambda: path_var.set(''),
                            bg=self.colors['white'], 
                            fg=self.colors['accent'],
                            font=('Segoe UI', 10, 'bold'),
                            width=3,
                            relief=tk.FLAT)
        clear_btn.pack(side=tk.LEFT, padx=(10, 0))
    
    def create_thresholds_section(self, parent):
        if IN_COLAB:
            return
            
        # Cadre pour les seuils
        frame = tk.Frame(parent, bg=self.colors['background'])
        frame.pack(fill=tk.X, pady=10)
        
        # Label
        label = tk.Label(frame, 
                        text="Seuils (1-199, séparés par des virgules)", 
                        font=('Segoe UI', 12), 
                        fg=self.colors['text'], 
                        bg=self.colors['background'])
        label.pack(anchor='w')
        
        # Entrée des seuils
        entry = tk.Entry(frame, 
                        textvariable=self.thresholds, 
                        font=('Segoe UI', 10), 
                        relief=tk.FLAT,
                        bg=self.colors['white'])
        entry.pack(fill=tk.X, pady=(5, 0))
    
    def select_scores_file(self):
        if IN_COLAB:
            return
            
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel des scores",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
        )
        if filepath:
            self.scores_path.set(filepath)
            self.prices_path.set(filepath)
    
    def validate_inputs(self, *args):
        if IN_COLAB:
            with self.output:
                self.output.clear_output()
                
                # Check if scores file was uploaded
                if not self.scores_path:
                    print("Erreur: Veuillez sélectionner le fichier des scores")
                    return None
                
                # Validate thresholds
                thresholds_str = self.thresholds_input.value.strip()
                if not thresholds_str:
                    print("Erreur: Veuillez entrer au moins un seuil")
                    return None
                
                try:
                    # Convert thresholds string to list of floats
                    thresholds = [float(x.strip()) for x in thresholds_str.split(',')]
                    
                    # Verify thresholds are between 1 and 199
                    if not all(1 <= x <= 199 for x in thresholds):
                        raise ValueError("Tous les seuils doivent être entre 1 et 199")
                    
                    # Create result dictionary
                    result = {
                        'prices_path': self.prices_path,
                        'scores_path': self.scores_path,
                        'thresholds': thresholds,
                        'file_content': self.uploaded_files.get(self.scores_path)
                    }
                    
                    # Store result
                    self.validation_result = result
                    
                    print("Succès: Tous les fichiers et seuils sont validés!")
                    
                    # Call callback if defined
                    if self.callback:
                        self.callback(result)
                    
                    return result
                    
                except ValueError as e:
                    print(f"Erreur: {str(e)}")
                    return None
        else:
            # Check if scores file was selected
            if not self.scores_path.get():
                messagebox.showerror("Erreur", "Veuillez sélectionner le fichier des scores")
                return None
            
            # Validation des seuils
            thresholds_str = self.thresholds.get().strip()
            if not thresholds_str:
                messagebox.showerror("Erreur", "Veuillez entrer au moins un seuil")
                return None
            
            try:
                thresholds = [float(x.strip()) for x in thresholds_str.split(',')]
                
                # Vérification que les seuils sont entre 1 et 199
                if not all(1 <= x <= 199 for x in thresholds):
                    raise ValueError("Tous les seuils doivent être entre 1 et 199")
                
                # Préparation du résultat
                result = {
                    'prices_path': self.prices_path.get(),
                    'scores_path': self.scores_path.get(),
                    'thresholds': thresholds
                }
                
                # Store result
                self.validation_result = result
                
                messagebox.showinfo("Succès", "Tous les fichiers et seuils sont validés!")
                
                # Fermer la fenêtre si un callback est défini
                if self.callback:
                    self.callback(result)
                
                self.root.quit()
                return result
                
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))
                return None
    
    def get_paths(self):
        return self.validation_result
    
    def run(self):
        # Create and display the UI
        self.create_ui()
        
        if not IN_COLAB:
            # Run the tkinter event loop
            self.root.mainloop()
        
        return self.validation_result

# Example usage function
def example_usage():
    def on_validate(data):
        print("Données validées:")
        print(f"- Seuils: {data['thresholds']}")
        print(f"- Fichier scores: {data['scores_path']}")
        
        # Example of reading the Excel file
        if IN_COLAB and 'file_content' in data and data['file_content'] is not None:
            # Save temporary file
            with open('temp_excel.xlsx', 'wb') as f:
                f.write(data['file_content'])
            
            # Read with pandas
            try:
                df = pd.read_excel('temp_excel.xlsx')
                print(f"Aperçu du fichier Excel:")
                if IN_COLAB:
                    display(df.head())
                else:
                    print(df.head())
            except Exception as e:
                print(f"Erreur lors de la lecture du fichier: {str(e)}")
            
            # Clean up
            if os.path.exists('temp_excel.xlsx'):
                os.remove('temp_excel.xlsx')
    
    # Create and run the selector
    app = ExcelFileSelector(callback=on_validate)
    return app.run()

if __name__ == "__main__":
    example_usage()