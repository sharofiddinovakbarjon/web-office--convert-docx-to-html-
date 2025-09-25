import React, { useEffect } from "react";
import { axiosAPI } from "./service/axiosAPI";
import { Editor } from "@tinymce/tinymce-react";

const App: React.FC = () => {
  const [templates, setTemplates] = React.useState<any[]>([]);
  const [editorValue, setEditorValue] = React.useState<string>("");

  const getTemplates = async () => {
    try {
      const response = await axiosAPI.get("/templates");
      if (response.status === 200) {
        setTemplates(response.data.templates);
      }
    } catch (error) {
      console.log(error);
    }
  };

  const getTemplateContent = async (templateName: string) => {
    try {
      const response = await axiosAPI.get(`/template/${templateName}`);

      if (response.status === 200) {
        setEditorValue(response.data);
      }
    } catch (error) {
      console.log(error);
    }
  };

  useEffect(() => {
    getTemplates();
  }, []);

  return (
    <div className="office-container">
      {/* Header with templates */}
      <div className="office-header">
        <div className="templates-section">
          <h3>Templates</h3>
          <div className="templates-list">
            {templates.map((template, index) => (
              <div key={index} className="template-item">
                <span>{template.name}</span>
                <button
                  className="select-btn"
                  onClick={() => getTemplateContent(template.name)}
                >
                  Select
                </button>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Main editor area */}
      <div className="office-main">
        <div className="paper-container">
          <div className="paper">
            <Editor
              apiKey="lkah965do05qqfb9852pfrb6itwnedze13xhj31j57h1ro9s"
              init={{
                plugins: [
                  // Core editing features
                  "anchor",
                  "autolink",
                  "charmap",
                  "codesample",
                  "emoticons",
                  "link",
                  "lists",
                  "media",
                  "searchreplace",
                  "table",
                  "visualblocks",
                  "wordcount",
                  // Your account includes a free trial of TinyMCE premium features
                  // Try the most popular premium features until Oct 3, 2025:
                  "checklist",
                  "mediaembed",
                  "casechange",
                  "formatpainter",
                  "pageembed",
                  "a11ychecker",
                  "tinymcespellchecker",
                  "permanentpen",
                  "powerpaste",
                  "advtable",
                  "advcode",
                  "advtemplate",
                  "ai",
                  "uploadcare",
                  "mentions",
                  "tinycomments",
                  "tableofcontents",
                  "footnotes",
                  "mergetags",
                  "autocorrect",
                  "typography",
                  "inlinecss",
                  "markdown",
                  "importword",
                  "exportword",
                  "exportpdf",
                ],
                toolbar:
                  "undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography uploadcare | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat",
                tinycomments_mode: "embedded",
                tinycomments_author: "Author name",
                mergetags_list: [
                  { value: "First.Name", title: "First Name" },
                  { value: "Email", title: "Email" },
                ],
                ai_request: (_request: any, respondWith: any) =>
                  respondWith.string(() =>
                    Promise.reject("See docs to implement AI Assistant")
                  ),
                uploadcare_public_key: "02ea2ba837042a9748ae",
                height: 1000,
                // menubar: false,
                // statusbar: false,
                content_style: `
                  body { 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    font-size: 11pt;
                    line-height: 1.15;
                    margin: 40px;
                    background: white;
                  }
                `,
              }}
              value={editorValue}
              onEditorChange={(content) => setEditorValue(content)}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
