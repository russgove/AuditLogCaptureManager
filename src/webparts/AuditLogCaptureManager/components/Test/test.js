import React, { useState } from "react";
import { useQuery } from "react-query";

export default function App() {
    const [searchTerm, setSearchTerm] = useState("");

    const { status, data } = useQuery(
        ["users", { status: "active", searchTerm }],
        (status, searchTerm) => {
            // return fetch(`/api/users/${status}?q=${searchTerm}`)
            console.log(status, searchTerm);
        }
    );

    return ( <
        input type = "search"
        value = { searchTerm }
        onChange = {
            (e) => setSearchTerm(e.target.value) }
        />
    );
}