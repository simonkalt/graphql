import JobList from "../components/JobList";
// import { jobs } from "../lib/fake-data";
// import { getJobs } from "../lib/graphql/queries";
import { useJobs } from "../lib/hooks";

// const jobs = await getJobs(); //.then((jobs) => console.log(jobs));

function HomePage() {
  const { jobs, loading, error } = useJobs();
  console.log("[HomePage] jobs:", jobs);
  if (loading) {
    return <h1>...Loading...</h1>;
  }
  if (error) {
    return <div className="has-text-danger">Data Unavailable</div>;
  }

  return (
    <div>
      <h1 className="title">Job Board</h1>
      <JobList jobs={jobs} />
    </div>
  );
}

export default HomePage;
