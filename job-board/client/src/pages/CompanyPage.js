import { useParams } from "react-router";
import { useState, useEffect } from "react";
// import { companies } from '../lib/fake-data';
import { getCompany } from "../lib/graphql/queries";
import JobList from "../components/JobList";

function CompanyPage() {
  const { companyId } = useParams();
  const [state, setState] = useState({
    company: null,
    loading: true,
    error: false,
  });

  useEffect(() => {
    // getCompany(companyId).then(setCompany);
    (async () => {
      try {
        const company = await getCompany(companyId);
        setState({ company, loading: false, error: false });
      } catch (error) {
        console.log("error:", JSON.stringify(error, null, 2));
        setState({ company: null, loading: false, error: true });
      }
    })();
  }, [companyId]);

  // const job = jobs.find((job) => job.id === jobId);
  console.log("[CompanyPage] state", state);
  const { company, loading, error } = state;

  if (loading) {
    return <h1>...Loading...</h1>;
  }
  if (error) {
    return (
      <div className="has-text-danger">
        <h1>...Error getting data, MoFo!</h1>
      </div>
    );
  }
  //const company = companies.find((company) => company.id === companyId);
  return (
    <div>
      <h1 className="title">{company.name}</h1>
      <div className="box">{company.description}</div>
      <h2 className="title is-5">Jobs at {company.name}</h2>
      <JobList jobs={company.jobs}></JobList>
    </div>
  );
}

export default CompanyPage;
