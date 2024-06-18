import React from "react";
import { useNavigate } from "react-router-dom";
import { ReactTyped } from "react-typed";
import { IoMdPlayCircle } from "react-icons/io";

function Hero() {
  const navigate = useNavigate();

  const handleGetStartedClick = () => {
    navigate("/file-dropper");
  };

  return (
    <div className="text-white">
      <div className="max-w-[800px] mt-[-96px] w-full h-screen mx-auto text-center flex flex-col justify-center">
        <p className="text-[#00df9a] font-bold p-1">With the Power of AI</p>
        <h1 className="md:text-7xl sm:text-6xl text-4xl font-bold md:py-6">
          Transform your CV
        </h1>
        <div className="flex justify-center items-center">
          <p className="md:text-5xl sm:text-4xl text-xl font-bold py-4">
            Don't waste your
          </p>
          <ReactTyped
            className="md:text-5xl sm:text-4xl text-xl font-bold md:pl-3 pl-2"
            strings={["Money", "Time", "Staff"]}
            typeSpeed={120}
            backSpeed={140}
            loop
          />
        </div>
        <div className="flex items-center mx-auto my-4 space-x-7">
          <button
            onClick={handleGetStartedClick}
            className="bg-[#00df9a] w-[200px] rounded-md font-medium my-6 py-3 hover:bg-[#00b87a] hover:scale-105 transition-transform duration-200"
          >
            Get Started
          </button>
          <button className="bg-white text-[#000300] w-[200px] rounded-md font-medium my-6 py-3 flex items-center justify-center hover:text-[#333333] hover:scale-105 transition-transform duration-200">
            <span>Watch Trailer</span>
            <IoMdPlayCircle size={20} className="ml-1" />
          </button>
        </div>
      </div>
    </div>
  );
}

export default Hero;
