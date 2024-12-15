import requests
import subprocess
import json
import os
import time
import pickle
import win32com.client
from pathlib import Path

# Prompt templates
PROMPTS = {
    "search_decision": """
    As a file search decision assistant, strictly return whether file search is needed in JSON format.

    User request: {user_request}

    Strictly return the following JSON format:
    {
        "needs_search": true/false,
        "search_keywords": "search keywords"
    }

    Notes:
    1. Only return JSON content, do not output any text other than JSON format!!!
    2. If the request clearly points to a file, set needs_search to true
    3. search_keywords should be the most critical filename or keyword
    4. If it's a command like creating or generating a file, no file search is needed
    5. if the command is to open, start, run, or copy a file should the file be searched and the search results returned; if it is for output, print, or view, no search is needed.
    6 .Only initiate a search when specific file names are involved.
    """,
    
    "command_generator": """
    As a Windows command generation assistant, generate precise commands based on user request and file search results.

    User request: {user_request}

    File search results:
    {search_results}

    Return JSON format:
    {
        "command": "specific Windows command",
        "description": "command description in English",
        "success": true/false
    }

    Requirements:
    1. Only return JSON, do not output any text other than JSON format!!!
    2. Command must accurately correspond to user intent
    3. If no valid command can be generated, set success to false
    4. The output command must strictly follow Windows command line syntax.
    5.It must strictly follow the syntax of Windows command line.

    
    
    Example:  
    Input "Create a folder named test3333 on C drive"  
    Return: {"command": "mkdir C:\test3333", 
            "description": "Create test3333 folder on C drive", 
            "success": true} 
    Example:  
    Input "What are the processes that are occupying 11434"  
    Return: {"command": "netstat -ano | findstr 11434", 
            "description": "check the processes that are occupying 11434", 
            "success": true} 
    Example:  
    Input "Copy the files from the 3d.model folder to the code6 folder"  
    Return: {"command": "copy "C:\zhubo\3d.model\*" "C:\code6" ", 
            "description": "Copy the files from the 3d.model folder to the code6 folde", 
            "success": true} 
    """ 
}


class CommandAssistant:
    def __init__(self):
        """Initialize command line assistant"""
        # Basic configuration
        self.ollama_base_url = "http://localhost:11434/api/generate"
        self.model = "llama3.1:8b"
        self.prompts = PROMPTS
        
        # Conversation history
        self.conversation_history = []
        self.max_history = 5
        self.history_file = "conversation_history.pkl"
        
        print("System: Initializing command assistant...")
        self.load_conversation_history()
        print("System: Ready")
        self.print_help()

    def print_help(self):
        """Print help information"""
        print("\nSupported Operations:")
        print("1. File Operations:")
        print("   - Open/Start/Run/View [filename]")
        print("2. Other Commands:")
        print("   - Create folder [name]")
        print("   - Copy file [source file] [destination]")
        print("   - Move file [source file] [destination]")
        print("   Other Windows command line operations")
        print("3. System Commands:")
        print("   - exit : Exit program")

    def load_conversation_history(self):
        """Load conversation history"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'rb') as f:
                    self.conversation_history = pickle.load(f)
                print(f"Loaded {len(self.conversation_history)} conversation records")
        except Exception as e:
            print(f"Failed to load history: {str(e)}")
            self.conversation_history = []
            
    def save_conversation_history(self):
        """Save conversation history"""
        try:
            with open(self.history_file, 'wb') as f:
                pickle.dump(self.conversation_history, f)
        except Exception as e:
            print(f"Failed to save history: {str(e)}")
            
    def add_to_history(self, user_input, ai_response):  
        """Add conversation to history"""  
        try:  
            ai_dict = json.loads(ai_response)  
            readable_response = ai_dict.get('description', str(ai_response))  
        except (json.JSONDecodeError, TypeError):  
            readable_response = str(ai_response)  

        self.conversation_history.append({  
            'user': user_input,  
            'ai': readable_response,  
            'timestamp': time.time()  
        })  
    
        if len(self.conversation_history) > self.max_history:  
            self.conversation_history = self.conversation_history[-self.max_history:]  
    
        self.save_conversation_history()  

    def quick_file_search(self, filename):  
        """Quickly search files using system API"""  
        results = []  
        print(f"\nSearching for '{filename}'...")  
    
        try:  
            drives = [d + ":\\" for d in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(d + ":")]  
            for drive in drives:  
                try:  
                    pattern = f"**/{filename}*"  
                    for path in Path(drive).glob(pattern):  
                        if path.is_dir() or path.is_file():  
                            stats = path.stat()  
                            results.append({  
                                'type': 'directory' if path.is_dir() else 'file',  
                                'name': path.name,  
                                'path': str(path.absolute()),  
                                'parent': str(path.parent),  
                                'size': stats.st_size if path.is_file() else 0,  
                                'modified': stats.st_mtime  
                            }) 
                              
                            print(f"Name: {result['name']}")  
                            print(f"Type: {result['type']}")  
                            print(f"Full Path: {result['path']}")  
                            print(f"Parent Directory: {result['parent']}")  
                            
                            if result['type'] == 'file':  
                                print(f"File Size: {self._format_size(result['size'])}")  
                                print(f"Last Modified: {time.ctime(result['modified'])}")  
                            
                            print("-" * 40)  
                except Exception as e:  
                      
                    continue  

            for drive in drives:  
                try:  
                    for root, dirs, files in os.walk(drive):  
                        search_items = dirs + files  
                        matches = [item for item in search_items if filename.lower() in item.lower()]  
                    
                        for match in matches:  
                            full_path = os.path.join(root, match)  
                            try:  
                                stats = os.stat(full_path)  
                                results.append({  
                                    'type': 'directory' if os.path.isdir(full_path) else 'file',  
                                    'name': match,  
                                    'path': full_path,  
                                    'parent': root,  
                                    'size': stats.st_size if os.path.isfile(full_path) else 0,  
                                    'modified': stats.st_mtime  
                                })  
                            except Exception:  
                                continue  
                except Exception as e:  
                    print(f"os.walk search error: {e}")  
                    continue  
                
        except Exception as e:  
            print(f"Search error: {e}")  
    
        unique_results = []  
        seen_paths = set()  
        for result in results:  
            if result['path'].lower() not in seen_paths:  
                seen_paths.add(result['path'].lower())  
                unique_results.append(result)  
    
        return unique_results
        
    def _format_size(self, size):
        """Format file size"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024:
                return f"{size:.2f} {unit}"
            size /= 1024
        return f"{size:.2f} PB"

    def determine_search_need(self, user_input):
        """Determine if file search is needed"""
        prompt = self.prompts["search_decision"].replace("{user_request}", user_input)
        response = self.call_model(prompt)
        
        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {"needs_search": False, "search_keywords": ""}

    def generate_command(self, user_input, search_results):  
        """Generate command based on search results"""  
        prompt = """  
    As a Windows command generation assistant, generate commands strictly following these requirements:  

    User request: {user_request}  

    File search results (if any):  
    {search_results}  

    Generation rules:  
    1. Carefully analyze user intent  
    2. Generate accurate Windows commands  
    3. Return JSON format:  
   {  
     "command": "specific cmd command",  
     "description": "command description in English",  
     "success": true/false  
   }  

Common command examples:  
- Create folder: mkdir C:\\test3333  
- Copy file: copy source.txt destination.txt  
- Move file: move source.txt destination.txt  

Special notes:  
- Use double backslash \\ for escaping  
- Paths must be accurate  
- If command cannot be generated, set success to false  
""".replace("{user_request}", user_input).replace("{search_results}", json.dumps(search_results))  
    
        response = self.call_model(prompt)  
        
        try:  
            return json.loads(response)  
        except json.JSONDecodeError:  
            try:  
                import re  
                json_match = re.search(r'\{.*\}', response, re.DOTALL)  
                if json_match:  
                    return json.loads(json_match.group(0))  
            except Exception as e:
                print(f"JSON parsing error: {e}")  
                 
            return {"command": "", "description": "Unable to generate command", "success": False}  

    def process_input(self, user_input):  
        """Process user input"""  
        if not user_input:  
            return  
        
        print(f"\nReceived user input: '{user_input}'")  
        search_results = []
        search_decision = self.determine_search_need(user_input)  
    
        if search_decision['needs_search']:  
            search_results = self.quick_file_search(search_decision['search_keywords'])  
        
            if not search_results:  
                print("No matching files found")  
                return  
        
            if len(search_results) > 1:  
                print("\nMultiple matching files found:")  
                for idx, result in enumerate(search_results, 1):  
                    print(f"{idx}. {result['name']} - {result['path']}")  
            
                try:  
                    choice = int(input("Select file (enter number): ")) - 1  
                    selected_result = [search_results[choice]]  
                except (ValueError, IndexError):  
                    print("Invalid selection")  
                    return  
            else:  
                selected_result = search_results  
        
            command_result = self.generate_command(user_input, selected_result)  
        else:  
            command_result = self.generate_command(user_input, [])  
    
        if command_result.get('success', False):  
            print(f"\nOperation description: {command_result['description']}")  
            print(f"Execute command: {command_result['command']}")  
        
            if input("Execute this command? (y/n): ").lower() == 'y':  
                try:  
                    subprocess.run(command_result['command'], shell=True)  
                    self.add_to_history(user_input, json.dumps(command_result))  
                except Exception as e:  
                    print(f"Command execution error: {e}")  
        else:
            print(search_results)  
            print("Unable to generate valid command")  
        
        
        
                 
    def call_model(self, prompt, context_length=5):  
        """  
        Call Ollama model with conversation history context  
        
        :param prompt: Current prompt  
        :param context_length: Number of historical conversations to include, default 5  
        :return: AI model response  
        """  
        try:  
            full_context = "Conversation History Context:\n"  
        
            recent_history = self.conversation_history[-context_length:]  
        
            for history_item in recent_history:  
                full_context += f"User: {history_item['user']}\n"  
                full_context += f"AI Assistant: {history_item['ai']}\n"  
                full_context += "---\n"  
        
            full_context += "\nCurrent Task Prompt:\n"  
            full_context += prompt  
        
            data = {  
                "model": self.model,  
                "prompt": full_context,  
                "stream": False,  
                "options": {  
                    "temperature": 0.5,  
                    "top_p": 0.9,        
                    "max_tokens": 2048   
                }  
            }  
        
            response = requests.post(self.ollama_base_url, json=data, timeout=30)  
        
            if response.status_code == 200:  
                result = response.json().get("response", "").strip()  
                return result  
            else:  
                print(f"AI service call failed: HTTP {response.status_code}")  
                return None  
    
        except requests.exceptions.Timeout:  
            print("AI service request timeout")  
            return None  
        except requests.exceptions.ConnectionError:  
            print("Unable to connect to AI service")  
            return None  
        except Exception as e:  
            print(f"AI call error: {str(e)}")  
            return None  

    def run(self):
        """Run command line interface"""
        while True:
            try:
                user_input = input("\nEnter command: ").strip()
                if user_input.lower() == 'exit':
                    self.save_conversation_history()
                    print("Program exited")
                    break
                if user_input:
                    self.process_input(user_input)
            except KeyboardInterrupt:
                self.save_conversation_history()
                print("\nProgram exited")
                break
            except Exception as e:
                print(f"Input processing error: {str(e)}")

def main():
    """Main function"""
    try:
        print("\n=== File System Command Assistant ===")
        print("Initializing, please wait...\n")
        app = CommandAssistant()
        app.run()
    except Exception as e:
        print(f"Program startup error: {str(e)}")

if __name__ == "__main__":
    main()