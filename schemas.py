from pydantic import BaseModel, Field
from typing import List, Optional

class PersonalInfo(BaseModel):
    full_name: str = Field(..., description="The candidate's full name (first and last name)")
    professional_title: str = Field(..., description="The candidate's job title or role, e.g., 'Senior Software Engineer'")
    email: str = Field(..., description="Candidate's contact email address")
    phone: str = Field(..., description="Candidate's contact phone number")
    location: str = Field(..., description="Candidate's city and state/country, e.g., 'San Francisco, CA'")
    linkedin: Optional[str] = Field(None, description="LinkedIn profile URL")
    github: Optional[str] = Field(None, description="GitHub profile URL")
    website: Optional[str] = Field(None, description="Personal website or portfolio URL")
    birth_info: Optional[str] = Field(None, description="Candidate's date and place of birth, e.g., '05/03/1984, Bokaro'")
    nationality: Optional[str] = Field(None, description="Candidate's nationality, e.g., 'Indian'")
    hobbies: Optional[str] = Field(None, description="Candidate's hobbies, e.g., 'Website Development, Learning new skills'")
    photo_base64: Optional[str] = Field(None, description="Base64 encoded profile photo string")
    power_statement: Optional[str] = Field(None, description="Candidate's personal tagline or power statement")


class WorkExperience(BaseModel):
    company: str = Field(..., description="Name of the company or organization")
    position: str = Field(..., description="Candidate's title/position, e.g., 'Senior Developer'")
    location: str = Field(..., description="City and state/country of the company")
    start_date: str = Field(..., description="Start date of employment, e.g., 'Jan 2021'")
    end_date: str = Field(..., description="End date of employment, e.g., 'Dec 2024' or 'Present'")
    description: List[str] = Field(..., description="Bullet points detailing achievements and responsibilities. Focus on quantifiable metrics and action verbs.")

class Education(BaseModel):
    institution: str = Field(..., description="Name of the school, college, or university")
    degree: str = Field(..., description="Degree or certification earned, e.g., 'B.S. in Computer Science'")
    location: str = Field(..., description="City and state/country of the institution")
    start_date: str = Field(..., description="Start date of education, e.g., 'Sep 2017'")
    end_date: str = Field(..., description="End date of education, e.g., 'May 2021' or 'Ongoing'")
    description: Optional[str] = Field(None, description="Additional details, e.g., GPA, honors, relevant coursework")

class SkillGroup(BaseModel):
    category: str = Field(..., description="Category of skills, e.g., 'Programming Languages', 'Frameworks', 'Soft Skills'")
    skills: List[str] = Field(..., description="List of specific skills within this category")

class Project(BaseModel):
    name: str = Field(..., description="Name of the project")
    role: Optional[str] = Field(None, description="Candidate's role in the project, e.g., 'Lead Developer'")
    link: Optional[str] = Field(None, description="Link to the repository or live site")
    description: str = Field(..., description="Brief description of the project and technologies used")

class Certification(BaseModel):
    name: str = Field(..., description="Name of the certification, e.g., 'AWS Certified Solutions Architect'")
    issuer: str = Field(..., description="Organization that issued the certification, e.g., 'Amazon Web Services'")
    date: str = Field(..., description="Date of completion or issuance, e.g., 'Mar 2024'")

class Language(BaseModel):
    name: str = Field(..., description="Language name, e.g., 'English', 'Spanish'")
    proficiency: str = Field(..., description="Candidate's proficiency level, e.g., 'Native', 'Fluent', 'Conversational'")

class ResumeData(BaseModel):
    personal_info: PersonalInfo = Field(..., description="Contact details")
    summary: str = Field(..., description="A professional summary (2-4 sentences) acting as a career pitch")
    experience: List[WorkExperience] = Field(default_factory=list, description="Work history in reverse-chronological order")
    education: List[Education] = Field(default_factory=list, description="Academic history")
    skills: List[SkillGroup] = Field(default_factory=list, description="Categorized list of skills")
    projects: List[Project] = Field(default_factory=list, description="Side projects or open source contributions")
    certifications: List[Certification] = Field(default_factory=list, description="Professional certifications")
    languages: List[Language] = Field(default_factory=list, description="Languages spoken")
    accomplishments: List[str] = Field(default_factory=list, description="Notable achievements and accomplishments")
