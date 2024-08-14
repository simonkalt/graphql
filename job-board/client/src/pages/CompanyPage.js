import { useParams } from "react-router";
// import { companies } from '../lib/fake-data';
import { companyByIdQuery } from "../lib/graphql/queries";
import JobList from "../components/JobList";
import { useQuery } from "@apollo/client";

function CompanyPage() {
  const { companyId } = useParams();
  const { data, error, loading } = useQuery(companyByIdQuery, {
    variables: { id: companyId },
  });

  console.log("[CompanyPage]", { data, error, loading });
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
  const { company } = data;

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
