import React, { useState, useEffect } from "react";

const emptyTemplates = {
  "Professional Summary": {
    "Professional Summary Text": "",
    "Professional Summary Bullet Points": ["", ""],
  },
  "Personal Information": {
    Firstname: "",
    Surname: "",
    Birthday: "",
    Nationality: "",
    "Marital Status": "",
    Availability: "",
    "Current Role": "",
    "Additional Information": "",
  },
  "Contact Information": {
    Address: "",
    "First Phone Number": "",
    "Second Phone Number": "",
    Email: "",
    "Additional Information": "",
  },
  Languages: [{ Name: "", Level: "" }],
  "Working Experience": [
    {
      Title: "",
      Location: "",
      Description: "",
      "Bullet Points": ["", ""],
      "Start Date": "",
      "End Date": "",
      Company: "",
      Website: "",
      "Additional Information": "",
    },
  ],
  Education: [
    {
      Diploma: "",
      Institution: "",
      "Start Date": "",
      "End Date": "",
      Grade: "",
      Location: "",
      Website: "",
      Description: "",
      "Bullet Points": ["", ""],
      "Additional Information": "",
    },
  ],
  Certificates: [
    {
      Title: "",
      "Start Date": "",
      "End Date": "",
      Institution: "",
      "Additional Information": "",
    },
  ],
  "Skills and Competencies": {
    Skills: [""],
    "Programming Languages": [{ Name: "", "Proficiency Level": "" }],
  },
  "Software and Technologies": [""],
  Hobbies: [""],
  "Additional Information": [
    {
      Title: "",
      "Start Date": "",
      "End Date": "",
      Description: "",
      Institution: "",
      Location: "",
      Address: "",
      Website: "",
      "Additional Information": "",
    },
  ],
};

const TextEditor = ({ formData, onSubmit }) => {
  const [data, setData] = useState(formData);

  useEffect(() => {
    setData(formData);
  }, [formData]);

  const handleInputChange = (path, value) => {
    const keys = path.replace(/\[(\d+)\]/g, ".$1").split(".");
    const newData = { ...data };

    let current = newData;
    keys.slice(0, -1).forEach((key) => {
      if (isNaN(parseInt(key, 10))) {
        if (!current[key]) current[key] = {};
        current = current[key];
      } else {
        if (!current[key]) current[key] = [];
        current = current[key];
      }
    });

    const lastKey = keys[keys.length - 1];
    current[lastKey] = value;
    setData(newData);
  };

  const addNewEntry = (path) => {
    const newData = { ...data };
    const keys = path.replace(/\[(\d+)\]/g, ".$1").split(".");

    let current = newData;
    keys.forEach((key) => {
      if (current[key] === undefined) {
        current[key] = Array.isArray(emptyTemplates[key]) ? [] : {};
      }
      current = current[key];
    });

    const lastKey = keys[keys.length - 1];
    const template = emptyTemplates[lastKey];

    if (Array.isArray(current)) {
      if (typeof template === "string") {
        current.push(template);
      } else {
        current.push({ ...template });
      }
    } else if (typeof current === "object") {
      let newKey = 1;
      while (`New Entry ${newKey}` in current) {
        newKey++;
      }
      current[`New Entry ${newKey}`] = { ...template };
    }

    setData(newData);
  };

  const renderInputField = (key, path, value) => (
    <div key={path}>
      <label className="text-md text-gray-300 font-medium">
        {key.replace(/_/g, " ")}
      </label>
      <input
        type="text"
        className="block w-full px-3 py-2 border border-gray-600 bg-gray-800 text-white rounded-md shadow-sm focus:outline-none focus:ring-[#00df9a] focus:border-[#00df9a]"
        value={value}
        onChange={(e) => handleInputChange(path, e.target.value)}
      />
    </div>
  );

  const renderArrayFields = (array, path) =>
    array.map((item, index) => {
      const itemPath = `${path}[${index}]`;
      if (typeof item === "string") {
        return renderInputField(`${path} ${index + 1}`, itemPath, item);
      } else if (typeof item === "object") {
        return (
          <div key={itemPath} className="mb-8">
            {Object.entries(item).map(([key, value]) =>
              renderInputField(key, `${itemPath}.${key}`, value)
            )}
          </div>
        );
      }
      return null;
    });

  const renderFormFields = (obj, path = "") =>
    Object.entries(obj).map(([key, value]) => {
      const newPath = path ? `${path}.${key}` : key;
      if (typeof value === "string") {
        return renderInputField(key, newPath, value);
      } else if (Array.isArray(value)) {
        return (
          <div key={newPath}>
            <h2 className="text-3xl text-[#00df9a] font-bold mt-4 mb-2">
              {key}
            </h2>
            {renderArrayFields(value, newPath)}
            {emptyTemplates[key] && (
              <button
                type="button"
                onClick={() => addNewEntry(newPath)}
                className="mt-2 px-3 py-1 bg-[#00df9a] text-white rounded hover:bg-[#00b87a] hover:scale-105 transition-transform duration-200"
              >
                Add New {key}
              </button>
            )}
          </div>
        );
      } else {
        return (
          <div key={newPath}>
            <h2 className="text-3xl text-[#00df9a] font-bold mt-4 mb-2">
              {key}
            </h2>
            {renderFormFields(value, newPath)}
            {emptyTemplates[key] && (
              <button
                type="button"
                onClick={() => addNewEntry(newPath)}
                className="mt-2 px-3 py-1 bg-[#00df9a] text-white rounded hover:bg-[#00b87a] hover:scale-105 transition-transform duration-200"
              >
                Add New Entry
              </button>
            )}
          </div>
        );
      }
    });

  const handleSubmit = (event) => {
    event.preventDefault();
    onSubmit(data); // Pass the edited data back to the parent component
  };

  return (
    <div className="min-h-screen p-5 bg-[#000300] shadow-md rounded-lg overflow-y-auto">
      <form onSubmit={handleSubmit}>
        {renderFormFields(data)}
        <button
          type="submit"
          className="mt-4 px-4 py-2 bg-[#00df9a] text-white rounded hover:bg-[#00b87a] hover:scale-105 transition-transform duration-200"
        >
          Submit
        </button>
      </form>
    </div>
  );
};

export default TextEditor;
