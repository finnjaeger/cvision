import React, { useState } from "react";
import TextEditor from "./TextEditor";

const FileDropper = () => {
  const [code, setCode] = useState("");
  const [setFiles] = useState([]);
  const [downloadLink, setDownloadLink] = useState("");
  const [message, setMessage] = useState(
    "Drag and drop PDF files here or click to select"
  );
  const [isUploading, setIsUploading] = useState(false); // New state to track uploading status
  const [resumeData, setResumeData] = useState(null); // State to hold resume data
  const [uploadId, setUploadId] = useState(null); // State to hold uploadId
  const [language, setLanguage] = useState("en"); // State to hold the selected language
  const [anonymize, setAnonymize] = useState(false); // State to hold the anonymize setting

  const preventDefaults = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDragOver = (e) => {
    preventDefaults(e);
    if (!isUploading) {
      setMessage("Drop the files here ...");
    }
  };

  const handleDragLeave = (e) => {
    preventDefaults(e);
    if (!isUploading) {
      setMessage("Drag and drop PDF files here or click to select");
    }
  };

  const handleDrop = (e) => {
    preventDefaults(e);
    if (!isUploading) {
      const newFiles = e.dataTransfer.files;
      setFiles([...newFiles]);
      handleFiles(newFiles);
    }
  };

  const handleFiles = (selectedFiles) => {
    setIsUploading(true); // Start uploading
    setMessage("Your CV is being uploaded...");
    Array.from(selectedFiles).forEach(uploadFile);
  };

  const uploadFile = (file) => {
    const formData = new FormData();
    formData.append("file", file, file.name);

    const requestOptions = {
      method: "POST",
      body: formData,
      headers: {
        "x-api-key": code,
        debug_mode: "false", // debugger
      },
    };

    fetch(
      "https://8bhp1g0nti.execute-api.eu-central-1.amazonaws.com/default/uploadCV",
      requestOptions
    )
      .then((response) => response.json())
      .then((data) => {
        setUploadId(data.uploadId); // Store the uploadId
        checkStatus(data.uploadId);
      })
      .catch((error) => {
        console.log("Upload error", error);
        setMessage("Upload error. Please try again.");
        setIsUploading(false); // Reset uploading status on error
      });
  };

  const checkStatus = (uploadId) => {
    setMessage("Checking CV processing status...");
    const intervalId = setInterval(() => {
      const requestOptions = {
        method: "GET",
        headers: {
          "x-api-key": code,
        },
      };
      fetch(
        `https://8bhp1g0nti.execute-api.eu-central-1.amazonaws.com/default/checkStatus?upload_id=${uploadId}`,
        requestOptions
      )
        .then((response) => response.json())
        .then((data) => {
          if (data.process_status === "ready_to_retrieve") {
            clearInterval(intervalId);
            setMessage("Generating Download link...");
            setResumeData(data.data); // Set the initial resume data
          } else if (data.process_status === "in_progress") {
            setMessage("Your CV is still being processed...");
          } else {
            clearInterval(intervalId);
            setMessage("Processing failed. Please try again.");
            setIsUploading(false); // Reset uploading status on failure
          }
        })
        .catch((error) => {
          clearInterval(intervalId);
          setMessage("Status check failed. Please try again.");
          setIsUploading(false); // Reset uploading status on error
        });
    }, 5000);
  };

  const createCV = async (uploadId, editedResumeData) => {
    const raw = JSON.stringify({
      upload_id: uploadId,
      language: language, // Language code selected by the user
      anonymize: anonymize, // Anonymize setting
      resume_data: editedResumeData,
    });
    const requestOptions = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": code,
      },
      body: raw,
      redirect: "follow",
    };

    try {
      const response = await fetch(
        "https://8bhp1g0nti.execute-api.eu-central-1.amazonaws.com/default/cv_creation",
        requestOptions
      );
      const result = await response.json();
      if (result && result.body) {
        const resultBody = JSON.parse(result.body);
        const downloadUrl = resultBody.resume;
        setIsUploading(false); // End uploading
        setDownloadLink(downloadUrl);
        setMessage("Your CV is ready to download.");
        setResumeData(null); // Reset to initial state
      } else {
        throw new Error("Failed to generate CV");
      }
    } catch (error) {
      console.error("CV Creation error:", error);
      setMessage("Error creating CV. Please try again.");
      setIsUploading(false); // Reset uploading status on error
    }
  };

  const handleEditorSubmit = (editedData) => {
    setResumeData(editedData); // Update the resume data
    createCV(uploadId, editedData); // Use the edited data to create CV
  };

  const toggleAnonymize = () => {
    setAnonymize((prev) => !prev);
  };

  return (
    <div className="bg-cover bg-center min-h-screen flex flex-col justify-center items-center px-4">
      {resumeData ? (
        <TextEditor formData={resumeData} onSubmit={handleEditorSubmit} />
      ) : (
        <>
          <div className="flex space-x-4 mb-4">
            <div className="flex items-center space-x-2">
              <div
                className={`w-6 h-6 rounded-full cursor-pointer ${
                  language === "en" ? "bg-[#00df9a]" : "bg-gray-400"
                }`}
                onClick={() => setLanguage("en")}
              />
              <span className="text-xl text-white">English</span>
            </div>
            <div className="flex items-center space-x-2">
              <div
                className={`w-6 h-6 rounded-full cursor-pointer ${
                  language === "de" ? "bg-[#00df9a]" : "bg-gray-400"
                }`}
                onClick={() => setLanguage("de")}
              />
              <span className="text-xl text-white">German</span>
            </div>
            <div className="flex items-center space-x-2">
              <div
                className={`w-6 h-6 rounded-full cursor-pointer ${
                  language === "fr" ? "bg-[#00df9a]" : "bg-gray-400"
                }`}
                onClick={() => setLanguage("fr")}
              />
              <span className="text-xl text-white">French</span>
            </div>
            <div className="flex items-center space-x-2">
              <div
                className={`w-6 h-6 rounded-full cursor-pointer ${
                  anonymize ? "bg-[#00df9a]" : "bg-red-500"
                }`}
                onClick={toggleAnonymize}
              />
              <span className="text-xl text-white">
                {anonymize ? "Anonymize" : "Anonymize"}
              </span>
            </div>
          </div>
          <input
            type="text"
            className="w-full max-w-md p-2 text-xl text-white bg-gray-700 rounded-lg shadow-md mb-4"
            placeholder="Enter your API key here"
            value={code}
            onChange={(e) => setCode(e.target.value)}
          />
          <div
            className={`border-2 border-transparent border-dashed rounded-xl p-20 w-full max-w-md text-center bg-white bg-opacity-10 hover:bg-opacity-20 hover:scale-110 transition duration-300 backdrop-blur-xl shadow-2xl ${
              isUploading ? "pointer-events-none" : ""
            }`}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onClick={() =>
              !isUploading && document.getElementById("fileElem").click()
            }
          >
            <p className="text-xl font-medium text-neutral-200 ">{message}</p>
            <input
              type="file"
              id="fileElem"
              multiple
              accept="application/pdf"
              style={{ display: "none" }}
              onChange={(e) => handleFiles(e.target.files)}
            />
          </div>
        </>
      )}
      {downloadLink && (
        <div className="mt-4 p-4 bg-white bg-opacity-80 rounded-lg shadow-xl text-center">
          <a
            className="text-lg font-medium underline hover:text-blue-600 transition duration-300 ease-in-out"
            href={downloadLink}
            download
          >
            Download
          </a>
        </div>
      )}
    </div>
  );
};

export default FileDropper;
