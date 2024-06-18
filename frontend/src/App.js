import React from 'react';
import { BrowserRouter as Router, Route, Routes, useLocation } from "react-router-dom";
import './index.css';
import Navbar from "./components/Navbar";
import Hero from "./components/Hero";
import FileDropper from "./components/FileDropper";
import TextEditor from "./components/TextEditor";

function App() {
  return (
    <Router>
      <PageLayout />
    </Router>
  );
}

function PageLayout() {
  const location = useLocation();

  return (
    <div className="App">
      {location.pathname !== "/file-dropper" && <Navbar />}
      <Routes>
        <Route path="/" element={<Hero />} />
        <Route path="/file-dropper" element={<FileDropper />} />
        <Route path="/text-editor" element={<TextEditor />} />
      </Routes>
    </div>
  );
}

export default App;
