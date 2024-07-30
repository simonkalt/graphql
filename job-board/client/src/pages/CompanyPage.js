import { useParams } from "react-router";
import { useState, useEffect } from "react";
// import { companies } from '../lib/fake-data';
import { getCompany } from "../lib/graphql/queries";

function CompanyPage() {
  const { companyId } = useParams();
  const [company, setCompany] = useState();
  useEffect(() => {
    getCompany(companyId).then(setCompany);
    console.log("[companId]:", companyId);
  }, [companyId]);

  // const job = jobs.find((job) => job.id === jobId);
  if (!company) {
    return <h1>...Loading...</h1>;
  }
  //const company = companies.find((company) => company.id === companyId);
  return (
    <div>
      <h1 className="title">{company.name}</h1>
      <div className="box">{company.description}</div>
      <h2 className="title is-5">Jobs at {company.name}</h2>
    </div>
  );
}

export default CompanyPage;
