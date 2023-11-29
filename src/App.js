import "./App.css";
import React, { useState } from "react";
import axios from "axios";
import { wait } from "@testing-library/user-event/dist/utils";

function App() {
  const [websiteUrl, setWebsiteUrl] = useState("");
  const [url2, setUrl2] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [results, setResults] = useState(null);
  const [error, setError] = useState(null);
  const [module_numbers, setModule_numbers] = useState([]);
  const [first, setFirst] = useState(false);
  const [second, setSecond] = useState(false);
  const [third, setThird] = useState(false);
  const [fourth, setFourth] = useState(false);

  const [smokeTest, setSmokeTest] = useState(false)

  const [userType, setUserType] = useState("");

  const handleInputChange = (event) => {
    setWebsiteUrl(event.target.value);
  };
  const userNameChange = (event) => {
    setUsername(event.target.value);
  };
  const passwordChange = (event) => {
    setPassword(event.target.value);
  };

  const allModuleSelector = () => {
    if (first && second && third && fourth) {
      setFirst(false);
      setSecond(false);
      setThird(false);
      setFourth(false);
      document.getElementById("Bank").checked = false;
      document.getElementById("Debtor").checked = false;
      document.getElementById("Creditor").checked = false;
      document.getElementById("Legal").checked = false;
    } else {
      setFirst(true);
      setSecond(true);
      setThird(true);
      setFourth(true);
      document.getElementById("Bank").checked = true;
      document.getElementById("Debtor").checked = true;
      document.getElementById("Creditor").checked = true;
      document.getElementById("Legal").checked = true;
    }
  };

  const firstStateChanger = () => {
    // console.log('first1',first);
    setFirst((prevState) => {
      return !prevState;
    });
    // console.log('first1',first);
  };
  const secondStateChanger = () => {
    // console.log('second',second);
    setSecond((prevState) => {
      return !prevState;
    });
    // console.log('second',second);
  };
  const thirdStateChanger = () => {
    // console.log('third',third);
    setThird((prevState) => {
      return !prevState;
    });
    // wait(5)
    // console.log('third',third);
  };
  const fourthStateChanger = () => {
    // console.log('fourth',fourth);
    setFourth((prevState) => {
      return !prevState;
    });
    // console.log('fourth',fourth);
  };

  const handleUserTypeChange = (event) => {
    setUserType(event.target.value);
  };

  const handleurl2change = (event) => {
    setUrl2(event.target.value);
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    setIsLoading(true);

    try {
      const response = await axios.post("http://localhost:5000/test", {
        website_url: websiteUrl,
        username: username,
        password: password,
        // module_numbers: module_numbers,
        first: first,
        second: second,
        third: third,
        fourth: fourth,
      });
      setResults(response.data);
      setError(null);
    } catch (error) {
      setResults(null);
      setError(error.message);
      console.log(error.message);
    } finally {
      setIsLoading(false);
      setModule_numbers([]);
    }
  };

  const handleSubmit2 = async (event) => {
    event.preventDefault();
    setIsLoading(true);
    try {
      const response = await axios.post(
        "http://localhost:5000/useraccessmatrix",
        {
          websiteUrl: websiteUrl,
          userType: userType,
          // username: username,
          // password: password,
          // module_numbers: module_numbers,
          // first: first,
          // second: second,
          // third: third,
          // fourth: fourth,
        }
      );
      setResults(response.data);
      setError(null);
    } catch (error) {
      setResults(null);
      setError(error.message);
      console.log(error.message);
    } finally {
      setIsLoading(false);
      // setModule_numbers([]);
    }
  };

  const handleSubmit3 = async (event) => {
    event.preventDefault();
    setIsLoading(true);
    try {
      const response = await axios.post(
        "http://localhost:5000/smoketest",
        {
          website_url: websiteUrl,
        username: username,
        password: password,
        // module_numbers: module_numbers,
        first: first,
        second: second,
        third: third,
        fourth: fourth,
        }
      );
      setResults(response.data);
      setError(null);
    } catch (error) {
      setResults(null);
      setError(error.message);
      console.log(error.message);
    } finally {
      setIsLoading(false);
      // setModule_numbers([]);
    }
  }

  return (
    <div>
      <form className="form" onSubmit={handleSubmit}>
        <div>
          <span style={{ "margin-bottom": "60px" }}>
            <label>URL</label>
            <input
              type="text"
              placeholder="Enter website URL"
              value={websiteUrl}
              onChange={handleInputChange}
              style={{ width: "80%" }}
            />
          </span>
          <span>
            <label>Username</label>
            <input
              type="text"
              placeholder="Enter Username"
              value={username}
              onChange={userNameChange}
            />
          </span>
          <span>
            <label>Password</label>
            <input
              type="text"
              placeholder="Enter Password"
              value={password}
              onChange={passwordChange}
            />
          </span>
          <div className="modules">
            <h1>Modules</h1>
            <section>
              <label htmlFor="Bank">Bank Confirmations</label>
              <input
                type="checkbox"
                name="Bank"
                id="Bank"
                onChange={firstStateChanger}
              />
            </section>
            <section>
              <label htmlFor="Debtor">Debtor Confirmations</label>
              <input
                type="checkbox"
                name="Debtor"
                id="Debtor"
                onChange={secondStateChanger}
              />
            </section>
            <section>
              <label htmlFor="Creditor">Creditor Confirmations</label>
              <input
                type="checkbox"
                name="Creditor"
                id="Creditor"
                onChange={thirdStateChanger}
              />
            </section>
            <section>
              <label htmlFor="Legal">Legal Matter Confirmations</label>
              <input
                type="checkbox"
                name="Legal"
                id="Legal"
                onChange={fourthStateChanger}
              />
            </section>
            <section>
              <label htmlFor="All">All</label>
              <input
                type="checkbox"
                name="All"
                id="All"
                onChange={allModuleSelector}
              />
            </section>
          </div>
        </div>
        <button style={{ marginBottom: "50px" }} className="button" type="submit" disabled={isLoading}>
          {isLoading ? "Loading..." : "Run Test"}
        </button>

        <button style={{ marginBottom: "50px" }} className="button" type="submit" disabled={isLoading} onClick={handleSubmit3}>
          {isLoading ? "Loading..." : "Run Smoke Test"}
        </button>

        <div>
          <span style={{height:"40px" }}>
            <label style={{ width: "40%" }}>URL</label>
            <input
              type="text"
              placeholder="Enter URL"
              value={url2}
              onChange={handleurl2change}
              style={{ width: "40%" }}
            />
          </span>
          <span style={{ height: "40px" }}>
            <label style={{ width: "40%" }}>User Access Matrix</label>
            <input
              type="number"
              placeholder="Enter User Type"
              value={userType}
              onChange={handleUserTypeChange}
              style={{ width: "40%" }}
            />
          </span>
          
        </div>

        <button
          onClick={handleSubmit2}
          className="button"
          type="submit"
          disabled={isLoading}
        >
          {isLoading ? "Loading..." : "Run Test"}
        </button>
      </form>

      {error && <p>Error: {error}</p>}

      {results && (
        <div>
          <p>Results:</p>
          <pre>{JSON.stringify(results, null, 2)}</pre>
        </div>
      )}
    </div>
  );
}

export default App;
